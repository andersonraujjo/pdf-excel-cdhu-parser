# CDHU-Parser-Precision 

O **CDHU-Parser-Precision** é um extrator de alta precisão desenvolvido em Python para converter planilhas de custos da CDHU (Companhia de Desenvolvimento Habitacional e Urbano) de PDF para Excel (`.xlsx`).

Diferente de conversores comuns, este projeto foi desenhado para superar as inconsistências de layout típicas dos documentos de engenharia civil, garantindo integridade total dos dados.

##  O Problema e a Solução Técnica (Pilar ADS)

Durante o desenvolvimento, identificamos que o maior "vilão" na extração de dados da CDHU é o deslocamento de caracteres. Unidades de medida (ex: `M3`, `M2`, `UN`) frequentemente se fundem ao texto da descrição do serviço, quebrando a estrutura de colunas.

### Diferencial: Tokenização Geométrica
Em vez de utilizar a estratégia padrão de extração de tabelas (que busca por linhas visíveis), este script utiliza **Tokenização Baseada em Coordenada Central (`xc`)**:
* **Mapeamento de Eixos:** O algoritmo extrai cada palavra individualmente e calcula seu centro geométrico.
* **Âncoras de Coluna:** Atribui cada token a uma coluna específica baseando-se em limites de pixels (eixo X) calibrados para o padrão CDHU.
* **Concatenação de Multilinhas:** Uma lógica de "olhar para trás" identifica descrições que ocupam várias linhas e as unifica em uma única célula no Excel.

##  Tecnologias
* **Python 3.10+**
* **pdfplumber:** Para extração granular e análise geométrica de tokens.
* **pandas:** Para estruturação de dados e exportação rápida.
* **CustomTkinter:** Interface moderna com suporte a processamento em segundo plano (Threading) para evitar travamentos da UI.
* **Openpyxl:** Formatação automática de larguras de coluna no arquivo final.

##  Instalação e Uso

1. **Requisitos:** Certifique-se de ter o Python instalado.
2. **Dependências:** Instale as bibliotecas necessárias:
   ```bash
   pip install -r requirements.txt
