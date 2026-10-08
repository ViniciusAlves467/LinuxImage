# ┌────────────────────────────────────────────────┐
# │  OUTLOOK / EMAIL - HTML para Power Automate      │
# └────────────────────────────────────────────────┘
# Substitui o bloco antigo (cor_prioridade + email_html montado na mão)
# dentro do `if total_alertas > 0:` da célula de notificações.
#
# No Databricks: cole o conteúdo de email_alertas.py numa célula anterior
# (ou use `%run ./email_alertas`) para ter `montar_email_html` disponível.

LINK_PAINEL = None  # Opcional: URL do dashboard "Gestão à Vista" (gera o botão no e-mail)

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

# --- Preview do email ---
displayHTML(email_html)
