# Automação de Contratos - Python

## Sobre o Projeto
Este projeto foi desenvolvido para automatizar a geração de contratos personalizados, otimizando processos que antes demandavam horas ou até dias de trabalho manual. Utilizando Python e o conceito de ambiente virtual (venv), o sistema garante isolamento de dependências e fácil replicação em diferentes ambientes.

A aplicação é estruturada com boas práticas de programação orientada a objetos, separando responsabilidades em classes e funções, o que facilita manutenção, evolução e entendimento do código por outros desenvolvedores.

## Principais Funcionalidades
Geração automática de três tipos de contratos:
- **Aditivo**: Para alterações contratuais.
- **Rescisão com aviso**: Para encerramento contratual com notificação.
- **Rescisão sem aviso**: Para encerramento contratual sem notificação.
Interface gráfica moderna e intuitiva (PyQt5), permitindo uso por qualquer pessoa, sem necessidade de conhecimento técnico.
Seleção de arquivos e pastas via diálogos interativos.
Barra de progresso para acompanhamento do processamento.
Conversão automática dos contratos gerados para PDF.
Estrutura modular, facilitando personalização e expansão.

## Como Funciona
O usuário alimenta duas planilhas CSV:
- **modelo_csv.csv**: Para contratos aditivos.
- **modelo_recisao.csv**: Para contratos de rescisão (com e sem aviso).
Ao iniciar o programa, uma janela é exibida solicitando a seleção do arquivo CSV e da pasta de destino dos contratos. O usuário escolhe o tipo de contrato desejado e, com apenas alguns cliques, o sistema processa os dados, gera os documentos personalizados e os salva automaticamente na pasta escolhida.

Esse fluxo elimina tarefas repetitivas e propensas a erro, garantindo padronização e agilidade.

## Requisitos do Sistema
- Python 3.8 ou superior
- Ambiente virtual (venv) recomendado
- Microsoft Word instalado (para conversão PDF via docx2pdf)
- Bibliotecas listadas em requirements.txt

## Instalação e Configuração
Clone o repositório:
```bash
git clone https://github.com/Jefferson170713/projeto-contrato-automatico.git
cd projeto-contrato-automatico
```

Crie e ative o ambiente virtual:
```bash
python -m venv venv
# No Windows:
venv\Scripts\activate
```

Instale as dependências:
```bash
pip install -r requirements.txt
```

## Execução
Para rodar o programa:
```bash
python script_aditivo.py
```

Para gerar um executável (Windows):
```bash
pyinstaller --onefile --windowed --add-data "Arquivos/hapvida_inside_circle.svg;Arquivos" script_aditivo.py
```

## Guia de Uso
Prepare as planilhas:
- Preencha modelo_csv.csv para contratos aditivos.
- Preencha modelo_recisao.csv para contratos de rescisão.
Siga o formato dos exemplos para garantir o correto funcionamento.

Execute o programa:
- Uma janela será aberta.
- Clique em "Selecionar arquivo CSV" e escolha a planilha desejada.
- Clique em "Selecionar pasta de saída" e escolha onde os contratos serão salvos.
- Marque o tipo de contrato a ser gerado (aditivo, rescisão com aviso, rescisão sem aviso).
- Clique para iniciar o processo.
- Acompanhe o progresso pela barra de progresso.
- Ao final, os arquivos estarão prontos na pasta escolhida.

## Exemplos de Aplicação
- Empresas que precisam gerar contratos personalizados para múltiplos clientes de forma rápida e padronizada.
- Departamentos jurídicos que desejam automatizar rotinas de rescisão contratual.
- Qualquer profissional que queira transformar dados de planilhas em documentos oficiais sem esforço manual.

## Estrutura do Projeto
- `script_aditivo.py`: Código principal da automação.
- `modelo_csv.csv`: Exemplo de planilha para contratos aditivos.
- `modelo_recisao.csv`: Exemplo de planilha para rescisões.
- `Arquivos/`: Modelos DOCX e ícones utilizados pelo sistema.
- `requirements.txt`: Lista de dependências do projeto.
- `LICENSE`: Licença MIT (com texto em português e inglês).
- `usar.txt`: Comandos úteis para ambiente virtual e execução.

## Boas Práticas e Dicas
- Sempre utilize o ambiente virtual para evitar conflitos de dependências.
- Mantenha os modelos DOCX atualizados conforme a necessidade do negócio.
- Valide os dados das planilhas antes de executar o programa para evitar erros.
- Consulte o arquivo usar.txt para comandos rápidos de configuração e execução.

## Licença
Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para detalhes em português e inglês.
