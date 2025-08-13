# Imovirtual → Apresentações Automáticas

Script em Python para extrair dados de anúncios do **Imovirtual** e gerar automaticamente apresentações em PowerPoint com o teu branding.

⚠️ **Aviso Legal**: Usa apenas dados e fotos com autorização do anunciante. O scraping pode violar os Termos de Uso do portal.

---

## Funcionalidades
- Lê lista de anúncios a partir de `urls.csv`
- Extrai:
  - Título
  - Preço
  - Localização
  - Área
  - Tipologia
  - Quartos
  - WCs
  - Descrição
  - Até 3 imagens (configurável)
- Gera:
  - `dados.csv` com todos os imóveis
  - `Apresentacoes_Imoveis.pptx` com 1 slide por imóvel + capa

---

## Instalação
```bash
pip install -r requirements.txt
playwright install
