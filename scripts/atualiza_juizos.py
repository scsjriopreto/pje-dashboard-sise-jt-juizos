#!/usr/bin/env python3
"""Gera o arquivo `juizos` a partir do cronograma de afastamentos (planilha no Google Drive).

Regras de vigência (horário de Brasília):
  * Ausência passa a valer às 10h do último dia útil ANTERIOR à data inicial.
  * Retorno passa a valer às 8h do primeiro dia útil POSTERIOR à data final.
  * Afastamentos encadeados (o retorno cairia no mesmo dia ou depois da saída
    seguinte) são tratados como um só período.

Redistribuição dos finais do juiz ausente (mesma vara):
  * vão para os titulares presentes; havendo mais de um, cada final vai para o
    titular que tem finais da mesma paridade (ex.: 9 -> ímpares, 0 -> pares);
  * sem titular presente, vão para quem estiver presente (ex.: simetria);
  * sem ninguém presente, vão para quem consta na coluna
    "Quem assina quando não há substituto" da planilha.

Uso:
  python scripts/atualiza_juizos.py [--xlsx arquivo.xlsx] [--agora "2026-12-07 08:00"] [--saida juizos]
Sem --xlsx, baixa a planilha do Drive (precisa estar compartilhada por link).
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
import unicodedata
import urllib.request
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

TZ = ZoneInfo("America/Sao_Paulo")
RAIZ = Path(__file__).resolve().parent.parent
DRIVE_ID = os.environ.get("CRONOGRAMA_DRIVE_ID", "1ypqbgBT9jV3UR43s7ZesXTQGgB6kNRdx")
HORA_SAIDA = time(10, 0)
HORA_RETORNO = time(8, 0)

avisos: list[str] = []


def normaliza(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return " ".join(txt.upper().split())


# ---------------------------------------------------------------- dias úteis
class Calendario:
    def __init__(self, cfg: dict):
        self.nao_uteis: set[date] = {date.fromisoformat(d) for d in cfg.get("datas", [])}
        for ini, fim in cfg.get("periodos", []):
            d, f = date.fromisoformat(ini), date.fromisoformat(fim)
            while d <= f:
                self.nao_uteis.add(d)
                d += timedelta(days=1)

    def util(self, d: date) -> bool:
        return d.weekday() < 5 and d not in self.nao_uteis

    def util_anterior(self, d: date) -> date:
        d -= timedelta(days=1)
        while not self.util(d):
            d -= timedelta(days=1)
        return d

    def util_posterior(self, d: date) -> date:
        d += timedelta(days=1)
        while not self.util(d):
            d += timedelta(days=1)
        return d


# ---------------------------------------------------------------- planilha
@dataclass
class Afastamento:
    ini: date
    fim: date
    tipo: str
    quem_assina: str | None


ROTULOS = {"FERIAS": "Férias", "COMPENSACAO": "Compensação", "LICENCA": "Licença"}


def rotulo(tipo: str) -> str:
    n = normaliza(tipo)
    for chave, r in ROTULOS.items():
        if n.startswith(chave):
            return r
    return str(tipo).strip().capitalize()


def para_data(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def baixa_planilha() -> bytes:
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_ID}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        dados = r.read()
    if not dados.startswith(b"PK"):
        raise SystemExit("Não foi possível baixar a planilha (verifique se o link está compartilhado).")
    return dados


def le_cronograma(dados: bytes) -> dict[str, list[Afastamento]]:
    wb = openpyxl.load_workbook(io.BytesIO(dados), data_only=True)
    ws = wb["Cronograma"] if "Cronograma" in wb.sheetnames else wb.worksheets[0]
    linhas = ws.iter_rows(values_only=True)
    cab = [normaliza(c or "") for c in next(linhas)]

    def col(prefixo):
        for i, c in enumerate(cab):
            if c.startswith(prefixo):
                return i
        raise SystemExit(f"Coluna '{prefixo}' não encontrada na aba Cronograma.")

    ci, cf, cj, ct, cq = col("DATA INICIAL"), col("DATA FINAL"), col("JUIZ"), col("TIPO"), col("QUEM ASSINA")
    por_juiz: dict[str, list[Afastamento]] = {}
    for n, row in enumerate(linhas, start=2):
        if not row or not row[cj]:
            continue
        ini, fim = para_data(row[ci]), para_data(row[cf])
        if not ini or not fim:
            avisos.append(f"Linha {n} ignorada: datas inválidas ({row[ci]!r}, {row[cf]!r}).")
            continue
        if fim < ini:
            avisos.append(f"Linha {n} ignorada: data final antes da inicial.")
            continue
        q = str(row[cq]).strip() if row[cq] else None
        por_juiz.setdefault(normaliza(row[cj]), []).append(Afastamento(ini, fim, str(row[ct] or ""), q))
    return por_juiz


# ---------------------------------------------------------------- períodos
def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


@dataclass
class Periodo:
    saida: datetime
    retorno: datetime
    trechos: list[list] = field(default_factory=list)  # [rotulo, ini, fim]
    quem_assina: list[str] = field(default_factory=list)

    def texto(self, hoje: date) -> str:
        # mostra só os trechos ainda não encerrados (ou o último, se todos já terminaram)
        trechos = [t for t in self.trechos if t[2] >= hoje] or self.trechos[-1:]
        partes = [f"{r} em {fmt(i)}" if i == f else f"{r} de {fmt(i)} a {fmt(f)}" for r, i, f in trechos]
        return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " e " + partes[-1]


def periodos_do_juiz(afs: list[Afastamento], cal: Calendario) -> list[Periodo]:
    res: list[Periodo] = []
    for a in sorted(afs, key=lambda a: (a.ini, a.fim)):
        saida = datetime.combine(cal.util_anterior(a.ini), HORA_SAIDA, TZ)
        retorno = datetime.combine(cal.util_posterior(a.fim), HORA_RETORNO, TZ)
        r = rotulo(a.tipo)
        if res and saida.date() <= res[-1].retorno.date():
            p = res[-1]
            p.retorno = max(p.retorno, retorno)
            ult = p.trechos[-1]
            if ult[0] == r and a.ini <= ult[2] + timedelta(days=1):
                ult[2] = max(ult[2], a.fim)
            else:
                p.trechos.append([r, a.ini, a.fim])
        else:
            p = Periodo(saida, retorno, [[r, a.ini, a.fim]])
            res.append(p)
        if a.quem_assina:
            p.quem_assina.append(a.quem_assina)
    return res


def ausencia_vigente(afs: list[Afastamento], cal: Calendario, agora: datetime) -> Periodo | None:
    for p in periodos_do_juiz(afs, cal):
        if p.saida <= agora < p.retorno:
            return p
    return None


# ---------------------------------------------------------------- montagem
def resolve_nome(apelido: str, todos: list[str]) -> str:
    """'Dra Adriana' -> 'ADRIANA FONSECA PERIN' (quando o primeiro nome é único)."""
    palavras = [w for w in normaliza(apelido).replace(".", " ").split() if w not in {"DR", "DRA"}]
    if not palavras:
        return apelido.strip()
    cand = [n for n in todos if all(w in n.split() for w in palavras)]
    if len(cand) == 1:
        return cand[0]
    avisos.append(f"Não identifiquei com certeza quem é '{apelido.strip()}'; usei o texto da planilha.")
    return normaliza(apelido)


def monta(base: dict, cronograma: dict[str, list[Afastamento]], cal: Calendario, agora: datetime) -> dict:
    todos = sorted({j["juiz"] for v in base["varas"].values() for j in v})
    for nome in cronograma:
        if nome not in todos:
            avisos.append(f"Juiz(a) '{nome}' está no cronograma mas não na divisão de finais (config/base.json).")

    vigentes = {n: ausencia_vigente(cronograma.get(n, []), cal, agora) for n in todos}
    saida: dict = {}
    for vara, juizes in base["varas"].items():
        presentes = [j for j in juizes if not vigentes[j["juiz"]]]
        titulares = [j for j in presentes if j["papel"] == "titular"]
        candidatos = titulares or presentes
        recebe = {j["juiz"]: list(j["finais"]) for j in presentes}
        sobra: list[int] = []
        for aus in (j for j in juizes if vigentes[j["juiz"]]):
            for f in aus["finais"]:
                if not candidatos:
                    sobra.append(f)
                    continue
                mesma_par = [c for c in candidatos if any(x % 2 == f % 2 for x in c["finais"])]
                recebe[(mesma_par or candidatos)[0]["juiz"]].append(f)

        bloco: dict = {}
        for j in juizes:
            p = vigentes[j["juiz"]]
            bloco[j["juiz"]] = [p.texto(agora.date())] if p else recebe[j["juiz"]]
        if sobra:
            quem = next((q for j in juizes if vigentes[j["juiz"]] for q in vigentes[j["juiz"]].quem_assina), None)
            if quem:
                nome = resolve_nome(quem, todos)
                bloco.setdefault(nome, [])
                if isinstance(bloco[nome], list) and all(isinstance(x, int) for x in bloco[nome]):
                    bloco[nome] = bloco[nome] + sobra
            else:
                avisos.append(f"{vara}: todos os juízes ausentes e ninguém na coluna 'Quem assina' — finais {sobra} sem responsável.")
        saida[vara] = bloco
    saida["Postos Avançados"] = base["postos_avancados"]
    return saida


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", help="planilha local (senão baixa do Drive)")
    ap.add_argument("--agora", help="simula data/hora, ex.: '2026-12-07 08:00'")
    ap.add_argument("--saida", default=str(RAIZ / "juizos"))
    args = ap.parse_args()

    agora = (datetime.strptime(args.agora, "%Y-%m-%d %H:%M").replace(tzinfo=TZ) if args.agora
             else datetime.now(TZ))
    base = json.loads((RAIZ / "config" / "base.json").read_text(encoding="utf-8"))
    cal = Calendario(json.loads((RAIZ / "config" / "feriados.json").read_text(encoding="utf-8")))
    dados = Path(args.xlsx).read_bytes() if args.xlsx else baixa_planilha()

    resultado = monta(base, le_cronograma(dados), cal, agora)
    texto = json.dumps(resultado, ensure_ascii=False, indent=2) + "\n"
    destino = Path(args.saida)
    anterior = destino.read_text(encoding="utf-8") if destino.exists() else None
    if texto != anterior:
        destino.write_text(texto, encoding="utf-8")
        print(f"[{agora:%d/%m/%Y %H:%M}] juizos atualizado.")
    else:
        print(f"[{agora:%d/%m/%Y %H:%M}] sem alterações.")
    for a in avisos:
        print("AVISO:", a)
    resumo = os.environ.get("GITHUB_STEP_SUMMARY")
    if resumo and avisos:
        with open(resumo, "a", encoding="utf-8") as fh:
            fh.write("### Avisos\n" + "".join(f"- {a}\n" for a in avisos))
    return 0


if __name__ == "__main__":
    sys.exit(main())
