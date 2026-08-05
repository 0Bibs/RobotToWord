# RobotToWord

Automatiza o protocolo de prints: pega um relatório em PDF do **Robot Structural
Analysis** e monta o `.docx` com um print por página, recortado e ajustado à
largura útil da folha — exatamente o que era feito à mão com a ferramenta de captura.

## O que ele faz

Para cada página do PDF:

1. rasteriza a página em alta resolução (200 DPI por padrão, contra os ~96 DPI de
   um print de tela — o resultado sai bem mais nítido);
2. recorta a moldura branca em volta do conteúdo;
3. insere a imagem no Word ocupando os **15,0 cm** de largura útil, sem distorcer;
4. começa cada print em uma página nova.

A folha sai igual à do modelo `Exemplo de forma para prints.docx`: A4 retrato,
margens de 3,0 cm nas laterais e 2,5 cm em cima e embaixo.

## Instalação

Precisa de Python 3.10 ou mais novo.

```bash
pip install -r requirements.txt
```

## Painel (para a equipe)

Quem não quiser linha de comando usa o painel:

```bash
python painel.py
```

No Windows, dê dois cliques em `painel.bat` — abre o painel sem a janela preta
de console atrás.

Ele tem três blocos: escolher os PDFs (avulsos ou uma pasta inteira), decidir se
sai um Word por PDF ou um só com tudo, e ajustar qualidade/recorte. A conversão
roda em segundo plano, com barra de progresso e um log do que está acontecendo;
no fim, o botão **Abrir pasta de saída** leva direto aos arquivos.

### Distribuir sem instalar Python

Para a equipe não precisar instalar nada, gere um executável único. Em uma
máquina Windows que tenha Python, rode:

```
build_exe.bat
```

Sai um `dist\RobotToWord.exe` (~50 MB) com Python e todas as bibliotecas
embutidos. Copie esse arquivo para uma pasta da rede e cada um usa com dois
cliques — sem instalação, sem PATH, sem permissão de administrador.

## Linha de comando

Um PDF, um Word ao lado dele:

```bash
python robot_to_word.py "Ligação 01 - Robot.pdf"
```

Uma pasta inteira de uma vez, cada PDF virando um `.docx` na pasta `saida/`:

```bash
python robot_to_word.py pasta_com_pdfs/ --out-dir saida/
```

Todas as ligações em um documento só, com o nome do arquivo como título de cada
trecho:

```bash
python robot_to_word.py pasta_com_pdfs/ --merge "Memorial de Ligações.docx" --title
```

## Opções da linha de comando

| Opção | Padrão | Para que serve |
|---|---|---|
| `--merge ARQUIVO.docx` | — | junta todos os PDFs em um único Word |
| `--out-dir PASTA` | ao lado do PDF | onde gravar os `.docx` |
| `--template MODELO.docx` | — | herda estilos, margens e cabeçalho/rodapé de um `.docx` seu |
| `--dpi N` | `200` | resolução da captura; 300 fica mais nítido e mais pesado |
| `--crop-mode MODO` | `uniform` | `uniform` recorta todas as páginas na mesma largura (texto do mesmo tamanho do começo ao fim); `tight` cola o recorte no conteúdo de cada página; `none` mantém as margens brancas do PDF |
| `--crop-threshold N` | `245` | quão claro um pixel precisa ser (0–255) para contar como fundo |
| `--crop-padding PT` | `4.0` | folga deixada em volta do conteúdo recortado |
| `--format png\|jpeg` | `png` | `jpeg` gera arquivo bem menor, com leve perda em textos finos |
| `--quality N` | `85` | qualidade do JPEG |
| `--no-page-break` | desligado | deixa os prints fluírem em vez de um por página |
| `--title` | desligado | escreve o nome do arquivo como título antes dos prints |

## Observações

- Páginas em branco são descartadas automaticamente.
- Um print mais alto que a área útil da folha é reduzido proporcionalmente para
  caber, em vez de ser cortado.
- Com `--template`, o conteúdo do modelo é esvaziado e só a formatação é
  aproveitada — o arquivo original não é alterado.
