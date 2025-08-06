# Gerenciador de Pastas G.A.A.L

## Descrição

Este projeto é um gerenciador de arquivos e pastas com interface gráfica, desenvolvido em Python utilizando PySide6. Permite organizar, copiar, mover, deletar e renomear arquivos e pastas de acordo com condições personalizadas, incluindo regras baseadas em planilhas Excel, estrutura física de diretórios e filtros por extensão/data. Suporta múltiplas origens, extração de arquivos ZIP, manutenção de hierarquia, criação de subpastas dinâmicas e exportação/importação de relatórios em Excel.

## Funcionalidades

- **Interface Gráfica Moderna**: Utiliza PySide6, com temas customizáveis e splash screen animada.
- **Múltiplas Origens**: Permite selecionar várias pastas ou arquivos ZIP como origem.
- **Filtros Avançados**: Filtra arquivos por extensão, data de modificação, nome, entre outros.
- **Condições Personalizadas**: Regras lógicas para processamento, podendo ser baseadas em colunas de Excel ou estrutura de subpastas.
- **Ações de Arquivo**: Copiar, mover, deletar, renomear arquivos e pastas.
- **Hierarquia e Subpastas**: Mantém ou recria hierarquia física, cria subpastas baseadas em condições.
- **Extração de ZIP**: Extrai arquivos ZIP automaticamente antes do processamento.
- **Relatórios**: Exporta e importa relatórios de execução em formato Excel.
- **Temas Customizáveis**: Vários temas visuais disponíveis e editor de temas integrado.
- **Cancelamento e Progresso**: Execução em thread com barra de progresso e opção de cancelamento.
- **Compactação Final**: Opcionalmente compacta a pasta de destino em um arquivo ZIP ao final.

## Estrutura do Projeto

```
├── executor.py         # Núcleo de processamento e lógica de condições
├── main.py             # Interface gráfica e integração com usuário
├── settings.json       # Configurações de temas e preferências
├── splash.png          # Imagem de splash
├── README.MD           # Documentação do projeto
```

## Principais Classes e Módulos

- **executor.py**
  - `FileManager`: Gerencia operações de arquivos e pastas.
  - `BooleanConditionEngine`, `ExcelConditionEngine`, `FolderConditionEngine`: Avaliam condições lógicas para processamento.
  - `Executor`: Orquestra o processamento em paralelo, aplica filtros e condições.
- **main.py**
  - `MainWindow`: Janela principal da interface gráfica.
  - `ExecutorThread`: Thread para execução do processamento com feedback ao usuário.
  - `ReportDialog`, `ThemeEditorDialog`, `ThemeManagerDialog`: Diálogos para relatório e edição de temas.
  - Funções utilitárias para splash, temas, flags e validação de configuração.

## Instalação

1. **Pré-requisitos**:
   - Python 3.10+ (recomenda-se 3.13)
   - Instalar dependências:
     ```powershell
     pip install PySide6 pandas openpyxl
     ```
2. **Executando**:
   - Via terminal:
     ```powershell
     python main.py
     ```

## Uso

1. Abra o aplicativo.
2. Selecione as pastas ou arquivos ZIP de origem.
3. Defina o destino e as opções desejadas (hierarquia, subpastas, filtros, etc).
4. Configure condições personalizadas (diretamente ou via importação de Excel).
5. Inicie o processamento e acompanhe o progresso.
6. Exporte o relatório de execução se necessário.

## Temas

O arquivo `settings.json` contém diversos temas visuais. É possível editar, criar ou remover temas diretamente pela interface.

## Relatórios

Após o processamento, é possível exportar o relatório em Excel, contendo informações sobre cada arquivo processado (origem, destino, ação).

## Contribuição

Contribuições são bem-vindas! Sinta-se livre para abrir issues ou pull requests.

## Autor

João Siqueira

---

Se desejar incluir instruções específicas de build, dependências adicionais ou exemplos de uso, basta solicitar!