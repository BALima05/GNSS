import subprocess
from pathlib import Path
from collections import Counter
import config
import utils
import re
import datetime


def extrair_ano_doy(nome_arq):
    """
    Extrai (ano, doy) do nome de um arquivo RINEX/produto IGS.
    Suporta: RINEX 3 (_YYYYDDD_), RINEX 2 (.YYo, .YYp, .YYn, .YYg) e arquivos BRDC (brdcDDD0.YYp).
    """
    match3 = re.search(r"_(\d{4})(\d{3})\d{4}_", nome_arq)
    if match3:
        return int(match3.group(1)), int(match3.group(2))

    match2 = re.search(r"(\d{3})[a-zA-Z0-9]\.(\d{2})[oOpPnNgG]$", nome_arq)
    if match2:
        return 2000 + int(match2.group(2)), int(match2.group(1))

    match_brdc = re.search(r"brdc(\d{3})\d\.\d{2}", nome_arq, re.IGNORECASE)
    if match_brdc:
        ano_m = re.search(r"\.(\d{2})", nome_arq)
        if ano_m:
            return 2000 + int(ano_m.group(1)), int(match_brdc.group(1))

    return None, None


def gps_date_converter(year, doy):
    full_year = 2000 + int(year) if int(year) < 100 else int(year)
    date_obj = datetime.datetime(full_year, 1, 1) + datetime.timedelta(days=int(doy) - 1)
    delta = date_obj - datetime.datetime(1980, 1, 6)
    return delta.days // 7, delta.days % 7, full_year, int(doy)


# GLONASS: número do canal de frequência (FCN)
# O gLAB precisa saber o FCN de cada satélite GLONASS. 
def ler_fcn_do_nav(nav_path):
    """Lê um BRDC RINEX 3 e devolve {'R01': fcn, ...}."""
    linhas = Path(nav_path).read_text(errors="ignore").splitlines()
    i = 0
    while i < len(linhas) and "END OF HEADER" not in linhas[i]:
        i += 1
    i += 1
    votos = {}
    while i < len(linhas):
        l = linhas[i]
        if re.match(r"R\d{2} ", l):
            if i + 2 < len(linhas):
                campo = linhas[i + 2][4 + 3 * 19: 4 + 4 * 19].replace("D", "E").strip()
                try:
                    votos.setdefault(l[:3], []).append(int(round(float(campo))))
                except ValueError:
                    pass
            i += 4
        else:
            i += 1
    return {s: Counter(v).most_common(1)[0][0] for s, v in votos.items()}


def montar_linhas_fcn(fcn):
    itens = [f"R{int(s[1:]):02d} {n:2d} " for s, n in sorted(fcn.items())]
    out = []
    for k in range(0, len(itens), 8):
        prefixo = f"{len(itens):3d} " if k == 0 else "    "
        out.append((prefixo + "".join(itens[k:k + 8])).ljust(60) + "GLONASS SLOT / FRQ #")
    return out


def inserir_fcn_no_obs(obs_in, obs_out, fcn):
    txt = Path(obs_in).read_text(errors="ignore")
    eol = "\r\n" if "\r\n" in txt else "\n"
    novas, em_header = [], True
    for l in txt.split(eol):
        if em_header and "GLONASS SLOT / FRQ #" in l:
            continue
        if em_header and "END OF HEADER" in l:
            novas.extend(montar_linhas_fcn(fcn))
            em_header = False
        novas.append(l)
    Path(obs_out).write_text(eol.join(novas))


# PPP COM gLAB
def processar_ppp_glab(arquivo_obs, pasta_produtos, pasta_nav, config_file, glab_path, pasta_saida):
    try:
        arquivo_obs = Path(arquivo_obs)
        pasta_produtos = Path(pasta_produtos)
        pasta_nav = Path(pasta_nav)
        pasta_saida = Path(pasta_saida)

        arquivo_pos = pasta_saida / arquivo_obs.with_suffix('.pos').name

        nav_files = list(pasta_nav.glob("*.[0-9][0-9][nNpPgG]")) + \
                    list(pasta_nav.glob("*.nav")) + list(pasta_nav.glob("*.NAV")) + \
                    list(pasta_nav.glob("*MN.rnx")) + list(pasta_nav.glob("brdc*"))
        arquivos_sp3 = list(pasta_produtos.glob("*.[sS][pP]3")) + list(pasta_produtos.glob("*.eph"))
        arquivos_clk = list(pasta_produtos.glob("*.[cC][lL][kK]"))

        ano, doy_int = extrair_ano_doy(arquivo_obs.name)
        if not (ano and doy_int):
            return f"⚠️ Pulei {arquivo_obs.name}: Não consegui identificar a data."

        wk, dw, _, _ = gps_date_converter(ano % 100, doy_int)
        nav_files = [f for f in nav_files if extrair_ano_doy(f.name) == (ano, doy_int)]
        padrao_curto = f"igs{wk}{dw}"
        arquivos_sp3 = [f for f in arquivos_sp3 if extrair_ano_doy(f.name) == (ano, doy_int) or padrao_curto in f.name.lower()]
        arquivos_clk = [f for f in arquivos_clk if extrair_ano_doy(f.name) == (ano, doy_int) or padrao_curto in f.name.lower()]

        if not arquivos_sp3: return f"⚠️ Pulei {arquivo_obs.name}: Faltam órbitas (.sp3)."
        if not arquivos_clk: return f"⚠️ Pulei {arquivo_obs.name}: Faltam relógios (.clk)."

        # --- GLONASS: gera cópia do obs com o FCN no cabeçalho ---
        obs_para_glab = arquivo_obs
        usa_glonass = arquivo_obs.name.startswith(("GLONASS_", "GPS_GLONASS_"))
        if usa_glonass:
            nav_brdc = [f for f in nav_files if f.name.upper().endswith("MN.RNX")]
            if not nav_brdc:
                return f"⚠️ Pulei {arquivo_obs.name}: Falta BRDC RINEX 3 (*MN.rnx) para obter o FCN do GLONASS."
            fcn = ler_fcn_do_nav(nav_brdc[0])
            if not fcn:
                return f"⚠️ Pulei {arquivo_obs.name}: Nenhum FCN GLONASS encontrado em {nav_brdc[0].name}."
            pasta_tmp = pasta_saida / "_obs_com_fcn"
            pasta_tmp.mkdir(exist_ok=True)
            obs_para_glab = pasta_tmp / arquivo_obs.name
            inserir_fcn_no_obs(arquivo_obs, obs_para_glab, fcn)

        # --- Comando: PPP = SOMENTE -input:orb e -input:clk (sem -input:nav) ---
        cmd = [
            str(glab_path),
            '-input:cfg', str(config_file),
            '-input:obs', str(obs_para_glab),
            '-output:file', str(arquivo_pos),
        ]
        for sp3 in arquivos_sp3:
            cmd.extend(['-input:orb', str(sp3)])
        for clk in arquivos_clk:
            cmd.extend(['-input:clk', str(clk)])

        result = subprocess.run(cmd, capture_output=True, text=True)

        # O gLAB escreve mensagens/erros DENTRO do arquivo -output:file
        msg_glab = arquivo_pos.read_text(errors="ignore") if arquivo_pos.exists() else ""
        erros = [l for l in msg_glab.splitlines() if l.startswith("ERROR")]
        n_out = sum(1 for l in msg_glab.splitlines() if l.startswith("OUTPUT"))

        if result.returncode != 0 or erros:
            return f"❌ gLAB falhou em {arquivo_obs.name} (code {result.returncode})\n   " + "\n   ".join(erros)
        if n_out == 0:
            return f"❌ {arquivo_obs.name}: gLAB rodou mas não gerou nenhuma linha OUTPUT (abra {arquivo_pos.name} e veja os INFO)."

        aviso = ""
        if usa_glonass:
            m = re.search(r"GLONASS FDMA frequencies disabled \(no frequency number\):\s*\S+\s+(.*)", msg_glab)
            if m and "R01-32" in m.group(1):
                aviso = "\n   ⚠️ gLAB desabilitou TODO o GLONASS (sem FCN) - resultado é só GPS!"
            elif m:
                aviso = f"\n   ℹ️ GLONASS sem FCN apenas em: {m.group(1).strip()} (satélites inativos, normal)"
        return f"✅ PPP: {arquivo_pos.name} ({n_out} épocas){aviso}"

    except Exception as e:
        return f"💥 Erro de execução em {arquivo_obs.name}: {e}"


def main():
    print("🌍 AUTOMAÇÃO DE PPP COM gLAB (Python Wrapper)")

    path_glab = Path(config.GLAB_PATH)
    if not path_glab.is_file():
        print(f"❌ Executável não encontrado: {path_glab}")
        return

    caminho_salvo = utils.carregar_estado("pasta_rinex_pronta")
    if not caminho_salvo:
        print("❌ Memória não encontrada. Execute o script 1 primeiro.")
        return

    path_rinex_obs = Path(caminho_salvo)
    base_dir = path_rinex_obs.parent

    path_nav = base_dir / "1.1 - Navegacao Broadcast"   # usada só para obter o FCN do GLONASS
    path_produtos = base_dir / "IGS_PRODUCTS"

    path_config = Path(config.CONFIG_FILE)
    if not path_config.is_file():
        print(f"❌ Arquivo de configuração não encontrado: {path_config}")
        return

    path_saida = base_dir / "RESULTADOS_PPP"
    path_saida.mkdir(parents=True, exist_ok=True)

    print(f"📁 Pasta de observação: {path_rinex_obs}")
    if not path_rinex_obs.exists():
        print(f"❌ A pasta {path_rinex_obs} não foi encontrada.")
        return
    if not path_nav.exists():
        print(f"⚠️ Pasta de navegação não encontrada em {path_nav}.")
        return

    print("\n🔍 Varrendo todas as subpastas (GPS, GLONASS, GPS_GLONASS)")
    arquivos_o = [
        a for a in path_rinex_obs.rglob("*")
        if a.is_file() and a.suffix.lower().endswith(("o", "rnx"))
    ]
    if not arquivos_o:
        print("Nenhum arquivo de observação encontrado.")
        return

    print(f"Iniciando PPP para {len(arquivos_o)} arquivos...")
    sucessos = 0
    for obs in sorted(arquivos_o):
        resultado = processar_ppp_glab(obs, path_produtos, path_nav, path_config, path_glab, path_saida)
        print(resultado)
        if resultado.startswith("✅"):
            sucessos += 1

    print(f"\n🏁 Finalizado: {sucessos}/{len(arquivos_o)} arquivos processados. Pasta: {path_saida}")
    if sucessos == 0:
        raise SystemExit(1)   


if __name__ == "__main__":
    main()