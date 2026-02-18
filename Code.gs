function setTelegramToken_() {
  const token = "PEGA_AQUI_TU_TOKEN_REAL";
  PropertiesService.getDocumentProperties().setProperty("TELEGRAM_BOT_TOKEN", token);
}
