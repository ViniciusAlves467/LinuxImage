import requests as req
import json
from datetime import datetime
from html import escape

# ══════════════════════════════════════════════════════════════
# NOTIFICAÇÕES - MICROSOFT TEAMS + OUTLOOK
# ══════════════════════════════════════════════════════════════
# Teams: Crie um Incoming Webhook no canal desejado:
# https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook
#
# Outlook: O corpo HTML abaixo pode ser usado com:
#   - Power Automate: trigger "When a row is added" na tabela Delta de alertas
#   - Ou: endpoint HTTP trigger do Power Automate chamado diretamente daqui
#
# E-mail com identidade visual do deck "Rota SU - Conecta 2026":
#   verde floresta dominante, rótulos em verde musgo, títulos em Georgia.
#   Layout 100% em <table> com estilos inline (compatível com Outlook desktop).
# ══════════════════════════════════════════════════════════════

# --- CONFIGURAÇÃO ---
TEAMS_WEBHOOK_URL       = "<COLE_URL_DO_WEBHOOK_TEAMS>"
POWER_AUTOMATE_HTTP_URL = "<COLE_URL_DO_TRIGGER_POWER_AUTOMATE>"  # Opcional
EMAIL_DESTINATARIOS     = ["gestor@usiminas.com"]  # Ajustar
LINK_PAINEL             = None  # Opcional: URL do dashboard "Gestão à Vista" (gera o botão no e-mail)
TABELA_ALERTAS          = None  # Opcional: tabela Delta de alertas (ex.: "catalogo.schema.iotracker_alertas"),
                                # usada só quando df_alertas não existe na sessão


# ┌────────────────────────────────────────────────┐
# │  TEMPLATE HTML DO E-MAIL (Rota SU)              │
# └────────────────────────────────────────────────┘
# --- Paleta (deck Rota SU) ---
VERDE_NOITE  = "#142B22"   # fundo do cabeçalho / rodapé
VERDE_FLORES = "#1B4332"   # painéis sobre o fundo escuro
VERDE_ACAO   = "#43A047"   # destaque principal
VERDE_MUSGO  = "#97BC62"   # eyebrow / rótulos
CINZA_SALVIA = "#C7D1CB"   # texto secundário no escuro
FUNDO_CLARO  = "#EEF2EF"   # fundo do corpo do e-mail
TEXTO        = "#1F2A24"
TEXTO_SUAVE  = "#5F6B65"
BORDA        = "#DDE5E0"

SERIF = "Georgia,'Times New Roman',serif"
SANS  = "Calibri,'Segoe UI',Arial,sans-serif"

# Nível do alerta -> (cor forte, fundo suave, rótulo)
ESTILO_NIVEL = {
    "Crítico": ("#C62828", "#FDECEC", "CRÍTICO"),
    "Atraso":       ("#E65100", "#FFF1E6", "ATRASO"),
    "Atenção": ("#B7791F", "#FFF8E1", "ATENÇÃO"),
}
ESTILO_PADRAO = (VERDE_ACAO, "#E4EFE6", "NO PRAZO")
ORDEM_PRIORIDADE = {"critica": 0, "alta": 1, "media": 2, "baixa": 3}


def _txt(valor, padrao="—"):
    """Converte para texto seguro em HTML (trata None/NaN)."""
    if valor is None or (isinstance(valor, float) and valor != valor):
        return padrao
    return escape(str(valor))


def _num(valor, padrao=0.0):
    try:
        v = float(valor)
        return padrao if v != v else v
    except (TypeError, ValueError):
        return padrao


def _kpi(valor, rotulo, cor):
    return f"""
      <td width="25%" valign="top" style="padding:0 6px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr><td bgcolor="{VERDE_FLORES}" align="center" style="background:{VERDE_FLORES}; border-radius:6px; padding:16px 8px 14px;">
            <div style="font-family:{SERIF}; font-size:30px; line-height:34px; font-weight:bold; color:{cor};">{valor}</div>
            <div style="font-family:{SANS}; font-size:11px; line-height:14px; color:{CINZA_SALVIA}; letter-spacing:1.5px; text-transform:uppercase; padding-top:4px;">{rotulo}</div>
          </td></tr>
        </table>
      </td>"""


def _barra(pct, cor):
    """Barra de progresso compatível com Outlook (duas células)."""
    pct = max(0, min(100, int(round(pct))))
    resto = 100 - pct
    cel_resto = (f'<td width="{resto}%" bgcolor="{BORDA}" style="background:{BORDA}; '
                 f'font-size:0; line-height:0;">&nbsp;</td>') if resto else ""
    cel_pct = (f'<td width="{pct}%" bgcolor="{cor}" style="background:{cor}; '
               f'font-size:0; line-height:0;">&nbsp;</td>') if pct else ""
    return f"""
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:6px;">
        <tr style="height:6px;">{cel_pct}{cel_resto}</tr>
      </table>"""


def _metrica(rotulo, valor, extra=""):
    return f"""
      <td width="25%" valign="top" style="padding:12px 14px; border-right:1px solid {BORDA};">
        <div style="font-family:{SANS}; font-size:10px; line-height:13px; color:{TEXTO_SUAVE}; letter-spacing:1.2px; text-transform:uppercase;">{rotulo}</div>
        <div style="font-family:{SANS}; font-size:16px; line-height:22px; color:{TEXTO}; font-weight:bold; padding-top:2px;">{valor}</div>
        {extra}
      </td>"""


def _card_rack(a):
    cor, fundo, rotulo = ESTILO_NIVEL.get(a.get("nivel_alerta"), ESTILO_PADRAO)
    pct_ciclo = _num(a.get("pct_ciclo"))
    score     = _num(a.get("score_pcp"))
    anomalias = _txt(a.get("anomalias"), "Nenhuma")

    if anomalias.strip().lower() in ("nenhuma", "", "—"):
        bloco_anomalia = f"""
          <td bgcolor="#F4F7F5" style="background:#F4F7F5; padding:10px 14px; border-radius:4px; font-family:{SANS}; font-size:13px; line-height:18px; color:{TEXTO_SUAVE};">
            &#10003;&nbsp; Nenhuma anomalia detectada
          </td>"""
    else:
        bloco_anomalia = f"""
          <td bgcolor="{fundo}" style="background:{fundo}; padding:10px 14px; border-radius:4px; font-family:{SANS}; font-size:13px; line-height:18px; color:{TEXTO};">
            <span style="color:{cor}; font-weight:bold;">&#9888;&nbsp; Anomalia:</span> {anomalias}
          </td>"""

    return f"""
    <tr><td style="padding:0 0 14px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
             bgcolor="#FFFFFF" style="background:#FFFFFF; border:1px solid {BORDA}; border-left:4px solid {cor}; border-radius:6px;">
        <!-- Cabeçalho do card -->
        <tr><td style="padding:16px 18px 12px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
            <td valign="middle">
              <div style="font-family:{SERIF}; font-size:20px; line-height:24px; font-weight:bold; color:{VERDE_NOITE};">{_txt(a.get('alias'))}</div>
              <div style="font-family:{SANS}; font-size:12px; line-height:16px; color:{TEXTO_SUAVE}; padding-top:2px;">
                Placa <b style="color:{TEXTO};">{_txt(a.get('placa'))}</b>
                &nbsp;&middot;&nbsp; Abastecimento <b style="color:{TEXTO};">{_txt(a.get('abastecimento'))}</b>
              </div>
            </td>
            <td valign="middle" align="right" style="white-space:nowrap;">
              <span style="display:inline-block; background:{cor}; color:#FFFFFF; font-family:{SANS}; font-size:11px; font-weight:bold; letter-spacing:1.5px; padding:5px 12px; border-radius:12px;">{rotulo}</span>
            </td>
          </tr></table>
        </td></tr>
        <!-- Métricas -->
        <tr><td style="padding:0 18px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
                 style="border:1px solid {BORDA}; border-radius:4px;"><tr>
            {_metrica("Local", _txt(a.get('localizacao')))}
            {_metrica("Dias no status", _txt(a.get('dias_no_status')))}
            {_metrica("% do ciclo", f"{pct_ciclo:.0f}%", _barra(pct_ciclo, cor))}
            <td width="25%" valign="top" style="padding:12px 14px;">
              <div style="font-family:{SANS}; font-size:10px; line-height:13px; color:{TEXTO_SUAVE}; letter-spacing:1.2px; text-transform:uppercase;">Score PCP</div>
              <div style="font-family:{SERIF}; font-size:18px; line-height:22px; color:{VERDE_ACAO}; font-weight:bold; padding-top:2px;">{score:.1f}</div>
              {_barra(score, VERDE_ACAO)}
            </td>
          </tr></table>
        </td></tr>
        <!-- Anomalia -->
        <tr><td style="padding:12px 18px 16px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>{bloco_anomalia}</tr></table>
        </td></tr>
      </table>
    </td></tr>"""


def montar_email_html(alertas, gerado_em=None, link_painel=None):
    """Monta o HTML do e-mail a partir de um DataFrame pandas de alertas."""
    gerado_em = gerado_em or datetime.now()
    alertas = alertas.copy()
    alertas["_ordem"] = alertas["prioridade"].map(ORDEM_PRIORIDADE).fillna(9)
    alertas = alertas.sort_values(["_ordem", "dias_no_status"], ascending=[True, False])

    total     = len(alertas)
    criticos  = int((alertas["nivel_alerta"] == "Crítico").sum())
    atrasos   = int((alertas["nivel_alerta"] == "Atraso").sum())
    atencao   = int((alertas["nivel_alerta"] == "Atenção").sum())

    cards = "".join(_card_rack(a) for a in alertas.to_dict("records"))

    botao = ""
    if link_painel:
        botao = f"""
        <tr><td align="center" style="padding:6px 0 26px;">
          <a href="{escape(link_painel)}" style="display:inline-block; background:{VERDE_ACAO}; color:#FFFFFF; font-family:{SANS}; font-size:14px; font-weight:bold; text-decoration:none; padding:12px 28px; border-radius:6px;">Abrir Gestão à Vista &rarr;</a>
        </td></tr>"""

    resumo = (f"{criticos} crítico(s), {atrasos} em atraso e {atencao} em atenção"
              if total else "Nenhum alerta ativo")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light only">
<title>IoTracker - Alertas de Racks</title>
</head>
<body style="margin:0; padding:0; background:{FUNDO_CLARO};">
<!-- Pré-cabeçalho (aparece na prévia da caixa de entrada) -->
<div style="display:none; max-height:0; overflow:hidden; mso-hide:all;">{total} rack(s) exigem ação: {resumo}.</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="{FUNDO_CLARO}" style="background:{FUNDO_CLARO};">
<tr><td align="center" style="padding:24px 12px;">
  <table role="presentation" width="680" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:680px;">

    <!-- ═══ Cabeçalho escuro ═══ -->
    <tr><td bgcolor="{VERDE_NOITE}" style="background:{VERDE_NOITE}; border-radius:10px 10px 0 0; padding:26px 30px 22px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td valign="middle">
          <div style="font-family:{SANS}; font-size:11px; line-height:14px; font-weight:bold; letter-spacing:3px; color:{VERDE_MUSGO};">GESTÃO À VISTA &nbsp;&middot;&nbsp; IOTRACKER</div>
          <div style="font-family:{SERIF}; font-size:28px; line-height:34px; font-weight:bold; color:#FFFFFF; padding-top:8px;">Alertas de Racks</div>
          <div style="font-family:{SANS}; font-size:14px; line-height:20px; color:{CINZA_SALVIA}; padding-top:4px;">{resumo}</div>
        </td>
        <td valign="top" align="right" style="white-space:nowrap;">
          <div style="font-family:{SANS}; font-size:12px; line-height:16px; color:{CINZA_SALVIA};">{gerado_em.strftime('%d/%m/%Y')}</div>
          <div style="font-family:{SANS}; font-size:12px; line-height:16px; color:{CINZA_SALVIA};">{gerado_em.strftime('%H:%M')}</div>
        </td>
      </tr></table>
    </td></tr>

    <!-- ═══ KPIs ═══ -->
    <tr><td bgcolor="{VERDE_NOITE}" style="background:{VERDE_NOITE}; padding:0 24px 26px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        {_kpi(total, "Alertas", "#FFFFFF")}
        {_kpi(criticos, "Críticos", "#EF5350")}
        {_kpi(atrasos, "Em atraso", "#FFA726")}
        {_kpi(atencao, "Atenção", "#FFD54F")}
      </tr></table>
    </td></tr>

    <!-- ═══ Corpo ═══ -->
    <tr><td bgcolor="{FUNDO_CLARO}" style="background:{FUNDO_CLARO}; padding:24px 0 6px;">
      <div style="font-family:{SANS}; font-size:11px; line-height:14px; font-weight:bold; letter-spacing:2.5px; color:{VERDE_ACAO}; padding:0 0 12px 2px;">RACKS QUE EXIGEM AÇÃO</div>
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
        {cards}
      </table>
    </td></tr>
    {botao}

    <!-- ═══ Rodapé escuro ═══ -->
    <tr><td bgcolor="{VERDE_NOITE}" style="background:{VERDE_NOITE}; border-radius:0 0 10px 10px; padding:18px 30px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td style="font-family:{SANS}; font-size:12px; line-height:18px; color:{CINZA_SALVIA};">
          <span style="font-family:{SERIF}; font-style:italic; color:{VERDE_MUSGO};">Rota SU</span>
          &nbsp;&bull;&nbsp; Rastreio, controle e gestão inteligente de racks
        </td>
        <td align="right" style="font-family:{SANS}; font-size:11px; line-height:16px; color:#7E8C85; white-space:nowrap;">
          Gerado automaticamente &middot; Databricks
        </td>
      </tr></table>
    </td></tr>

  </table>
</td></tr>
</table>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════
# EXECUÇÃO
# ══════════════════════════════════════════════════════════════

# --- Garantir que df_alertas existe (vem das células anteriores do notebook) ---
try:
    df_alertas
except NameError:
    if TABELA_ALERTAS:
        print(f"df_alertas n\u00e3o encontrado na sess\u00e3o - lendo da tabela {TABELA_ALERTAS}")
        df_alertas = spark.table(TABELA_ALERTAS)
    else:
        raise NameError(
            "df_alertas n\u00e3o existe nesta sess\u00e3o. Execute as c\u00e9lulas anteriores do notebook "
            "(menu da c\u00e9lula > 'Run all above') ou preencha TABELA_ALERTAS com a tabela Delta de alertas."
        )

# --- Filtrar alertas que precisam de notificação ---
df_alertas_pd = df_alertas.toPandas()
alertas_notificar = df_alertas_pd[df_alertas_pd["prioridade"].isin(["media", "alta", "critica"])]
total_alertas = len(alertas_notificar)

print(f"Alertas para notificar: {total_alertas}")
print(f"Total geral: {len(df_alertas_pd)} | Baixa: {len(df_alertas_pd[df_alertas_pd['prioridade']=='baixa'])}")
print("=" * 70)

if total_alertas > 0:
    # ┌────────────────────────────────────────────────┐
    # │  MICROSOFT TEAMS - Adaptive Card               │
    # └────────────────────────────────────────────────┘
    icone_nivel = {"Atenção": "⚠️", "Atraso": "\U0001f6a8", "Crítico": "\U0001f525"}

    card_body = [
        {
            "type": "TextBlock",
            "text": f"IoTracker - {total_alertas} Alerta(s) de Rack",
            "size": "Large", "weight": "Bolder", "color": "Attention"
        },
        {
            "type": "TextBlock",
            "text": f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            "size": "Small", "isSubtle": True
        },
    ]

    for _, a in alertas_notificar.iterrows():
        icone = icone_nivel.get(a["nivel_alerta"], "")
        card_body.append({"type": "ColumnSet", "columns": [
            {"type": "Column", "width": "auto", "items": [
                {"type": "TextBlock", "text": f"{icone} **{a['alias']}** ({a['placa']})",
                 "weight": "Bolder"}
            ]}
        ]})
        card_body.append({"type": "FactSet", "facts": [
            {"title": "Nível",        "value": a["nivel_alerta"]},
            {"title": "Localização",  "value": a["localizacao"]},
            {"title": "Dias",         "value": str(a["dias_no_status"])},
            {"title": "% Ciclo",      "value": f"{a['pct_ciclo']}%"},
            {"title": "Anomalias",    "value": a["anomalias"]},
            {"title": "Score PCP",    "value": str(a["score_pcp"])},
            {"title": "Abastecimento","value": a["abastecimento"]},
        ]})
        card_body.append({"type": "TextBlock", "text": "---", "spacing": "Small"})

    adaptive_card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": card_body
            }
        }]
    }

    # --- Enviar para Teams ---
    if TEAMS_WEBHOOK_URL != "<COLE_URL_DO_WEBHOOK_TEAMS>":
        try:
            resp = req.post(TEAMS_WEBHOOK_URL, json=adaptive_card, timeout=30)
            if resp.status_code in (200, 202):
                print(f"✅ Teams: Notificação enviada com sucesso!")
            else:
                print(f"❌ Teams: Erro {resp.status_code} - {resp.text[:200]}")
        except Exception as e:
            print(f"❌ Teams: Falha na conexão - {e}")
    else:
        print("⚠️  Teams: Configure TEAMS_WEBHOOK_URL para ativar")
        print("    Passos: Teams > Canal > Conectores > Incoming Webhook > Copiar URL")

    # ┌────────────────────────────────────────────────┐
    # │  OUTLOOK / EMAIL - HTML para Power Automate      │
    # └────────────────────────────────────────────────┘
    email_html = montar_email_html(alertas_notificar, link_painel=LINK_PAINEL)

    # --- Enviar para Power Automate (trigger HTTP) ---
    if POWER_AUTOMATE_HTTP_URL != "<COLE_URL_DO_TRIGGER_POWER_AUTOMATE>":
        n_criticos = int((alertas_notificar["prioridade"] == "critica").sum())
        payload = {
            "assunto": (f"\U0001f534 IoTracker | {n_criticos} crítico(s) - {total_alertas} rack(s) exigem ação"
                        if n_criticos else f"IoTracker | {total_alertas} rack(s) exigem ação"),
            "corpo_html": email_html,
            "destinatarios": EMAIL_DESTINATARIOS,
            "prioridade": "alta" if n_criticos else "normal",
        }
        try:
            resp = req.post(POWER_AUTOMATE_HTTP_URL, json=payload, timeout=30)
            print(f"✅ Power Automate: Trigger enviado ({resp.status_code})")
        except Exception as e:
            print(f"❌ Power Automate: Falha - {e}")
    else:
        print("\n⚠️  Outlook/Email: Configure POWER_AUTOMATE_HTTP_URL para ativar")
        print("    Passos no Power Automate:")
        print("    1. Novo fluxo > Trigger 'When an HTTP request is received'")
        print("    2. Ação: 'Send an email (V2)' do Outlook")
        print("    3. Usar corpo_html do payload como Body do email")
        print("    4. Copiar a URL do trigger e colar em POWER_AUTOMATE_HTTP_URL")

    # --- Preview do email ---
    print("\n" + "=" * 70)
    print("  PREVIEW DO EMAIL")
    print("=" * 70)
    displayHTML(email_html)

else:
    print("✅ Todos os racks estão no prazo. Nenhuma notificação necessária.")
