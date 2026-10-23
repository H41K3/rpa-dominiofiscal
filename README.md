# 🤖 RPA Automático - Cadastro de Acumuladores (Domínio Sistemas)

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)

## 📌 Sobre o Projeto
Este projeto é um **Robô de Automação de Processos (RPA)** desenvolvido em Python, focado em eliminar o trabalho manual e repetitivo na área fiscal/contábil. O script automatiza o preenchimento de centenas de registros complexos ("Acumuladores") dentro do ERP **Domínio Escrita Fiscal**, interagindo diretamente com a interface do sistema operacional (Desktop GUI).

A solução foi desenhada para atuar como um "funcionário digital", lendo dados brutos de planilhas Excel, higienizando as informações e injetando-as no sistema legado em alta velocidade e com 100% de precisão.

## 💼 Impacto no Negócio (Business Value)
- **Eliminação de Erros Humanos:** Garante que CFOPs, contas contábeis e impostos sejam preenchidos rigorosamente conforme a matriz aprovada no Excel.
- **Redução de Custo/Tempo:** Um processo de digitação que levaria dias foi reduzido a horas de processamento em background, libertando a equipa para tarefas analíticas.
- **Resiliência:** O sistema é imune a quedas de energia ou interrupções, possuindo um sistema de *checkpoint* que continua exatamente de onde parou.

## 🛠️ Destaques Técnicos e Arquitetura
Como o sistema legado não possui API aberta, a automação foi construída baseada na simulação de interação humana nativa via SO:

1. **ETL Híbrido (`pandas`):** Leitura de planilhas dinâmicas, formatação de tipos de dados (remoção de NaN, conversão de float para int) e mapeamento dinâmico de mais de 25 colunas booleanas e de valores.
2. **Navegação Cega Avançada (`pywinauto`):** Interação em alta cadência utilizando algoritmos de navegação via teclado (`SendKeys`), evitando a instabilidade comum da automação baseada em leitura de pixels (mouse click).
3. **Sistema de Checkpoint Contínuo:** Utilização de logs em tempo real. O robô regista o estado de sucesso de cada iteração. Ao reiniciar, o script cruza o Dataframe atual com o Log e processa apenas a diferença matemática (Delta).
4. **Kill Switch de Segurança (Interrupção via Hardware):** Implementação de uma trava de emergência utilizando `ctypes.windll`. O robô monitora de forma assíncrona a porta do teclado (tecla `ESC`) a cada 0.1 segundos, permitindo a finalização instantânea do processo (graceful shutdown) caso o utilizador necessite de recuperar o controlo físico da máquina.

## 🚀 Como Funciona a Lógica
1. Varredura do diretório pelo arquivo .xlsx mais recente.
2. Tratamento e exportação de uma cópia estruturada dos dados.
3. Verificação do arquivo de log local para descartar acumuladores já inseridos.
4. Contagem regressiva de segurança para focar a janela do ERP.
5. Loop de digitação iterando pelas abas do software (Geral, Notas, Impostos, CFOP, Contabilidade) através de atalhos nativos (`TAB`, `SHIFT+TAB`, `SPACE`).
6. Gravação e atualização do log.

## ⚙️ Tecnologias Utilizadas
- **Python 3.x**
- **Pandas:** Para manipulação, limpeza e filtro dos dados estruturados.
- **PyWinAuto:** Para injeção de comandos de teclado nativos no Windows.
- **Tkinter:** Utilizado de forma "headless" para manipular e higienizar a área de transferência do sistema (Clipboard).
- **Ctypes:** Acesso direto à API do Windows para monitorização do teclado físico.