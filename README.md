# 🗓️ Gerador Avançado de Escala de Voluntários

Uma aplicação web desenvolvida em Python com **Streamlit** para automação e gerenciamento inteligente de escalas de voluntários. O sistema realiza sorteios equilibrados, respeita metas específicas por reuniões e permite travar o número exato de participações por pessoa.

---

## ✨ Funcionalidades Principais

*   **Persistência de Dados (`Session State`):** Os resultados gerados ficam travados na tela. Você pode copiar o texto ou baixar os arquivos (Excel/TXT) sem medo de a escala sumir ou mudar de layout.
*   **Controle de Cota Individual:** Defina limites exatos de participação adicionando o número entre parênteses ao lado do nome (ex: `Tiago Chaves(3)` servirá exatamente 3 vezes). Pessoas sem parênteses entram na rotação de equilíbrio comum.
*   **Ajuste Dinâmico de Metas:** Uma tabela interativa permite definir quantas pessoas são necessárias para cada horário (`9:30h`, `11:30h` e `17:30h`) em cada domingo do mês de forma individual.
*   **Dupla Exportação:**
    *   **Excel (.xlsx) Profissional:** Gerado com design limpo usando `openpyxl`, contendo uma aba visual da escala, uma aba de estatísticas de participação e o texto formatado para cópia.
    *   **Texto Formatado (.txt):** Pronto para copiar e colar diretamente no WhatsApp ou grupos de aviso.

---

## 🚀 Como Executar o Projeto

### Pré-requisitos

Antes de começar, você vai precisar ter o **Python 3.8+** instalado em sua máquina.

### 1. Clonar ou criar a pasta do projeto
Crie uma pasta para o seu projeto e salve o arquivo do código como `app.py`.

### 2. Instalar as Dependências

Abra o seu terminal na pasta do projeto e instale as bibliotecas necessárias executando:

```bash
pip install streamlit pandas openpyxl