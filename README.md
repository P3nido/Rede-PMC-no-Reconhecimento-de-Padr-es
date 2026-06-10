# Rede PMC no Reconhecimento de Padrões

Este repositório contém a implementação de uma **Rede Neural Artificial do tipo Perceptron Multicamadas (PMC)** — também conhecida como *Multilayer Perceptron (MLP)* — aplicada a problemas de **Reconhecimento de Padrões**. O projeto abrange desde o treinamento do modelo com conjuntos de dados específicos até a etapa de validação e extração de métricas de desempenho.

## 📂 Estrutura do Projeto

Abaixo está a descrição dos principais arquivos e diretórios que compõem este repositório:

- **`main.py`**: Script principal em Python que contém a lógica de execução. Responsável por carregar os dados, treinar a rede PMC, realizar predições e gerar os resultados.
- **`requirements.txt`**: Lista de dependências e bibliotecas Python (como NumPy, Pandas, Matplotlib, Scikit-Learn, etc.) necessárias para rodar a aplicação.
- **`graphics/`**: Diretório destinado ao armazenamento das visualizações e gráficos gerados (ex: curvas de aprendizado, dispersão dos dados, fronteiras de decisão).

### 🗂️ Dados Utilizados
- `PP04_dados_treinamento.txt` / `.xls`: Conjunto de dados utilizado para a fase de **treinamento** da rede neural.
- `PP04_dados_validacao.txt` / `.xls`: Conjunto de dados não vistos pela rede, utilizado para a fase de **validação e teste**.

### 📊 Métricas e Resultados
- `Etapa5_metricas.xlsx`: Planilha consolidada com os resultados detalhados e as métricas de desempenho do modelo após as baterias de testes.
- `Tabela1_validacao.xlsx`: Tabela focada na análise da etapa de validação.

### 📚 Documentação Base
- `4_PMC.pdf`, `5_PMC-Aplicacoes.pdf`, `PP04_PMC-reconhecimento.pdf`: Guias, materiais teóricos e especificações das práticas/projetos originais, detalhando o funcionamento das Redes PMC e sua aplicação em padrões.

---

## 🚀 Como Executar

### Pré-requisitos
Certifique-se de ter o [Python 3.x](https://www.python.org/downloads/) instalado na sua máquina. É altamente recomendável a utilização de um ambiente virtual (como `venv` ou `conda`) para evitar conflitos de versões.

### Passos para a execução

1. **Clone este repositório:**
   ```bash
   git clone [https://github.com/P3nido/Rede-PMC-no-Reconhecimento-de-Padr-es.git](https://github.com/P3nido/Rede-PMC-no-Reconhecimento-de-Padr-es.git)
