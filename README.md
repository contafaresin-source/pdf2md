# PDF → Markdown (uso local)

App web local (Flask) para converter PDFs técnicos/científicos em Markdown e reduzir tokens na análise por IA.
Motores: **Docling**, **Marker** e **Microsoft Word + Pandoc**. Drag & drop de vários PDFs, fila, log de erros.

## Passo a passo no VS Code

1. Descompacte a pasta e abra no VS Code: `File > Open Folder > pdf2md`.
2. Abra o terminal integrado (`Ctrl+'`) e rode o setup:
   ```
   .\setup.bat            # base + pergunta se instala Docling/Marker
   .\setup.bat -All       # instala tudo sem perguntar
   .\setup.bat -Docling   # só o Docling (idem -Marker)
   ```
   - Usa o `uv`: baixa sozinho o Python 3.12 (Docling/Marker ainda não funcionam no 3.13/3.14) e recria ambientes antigos incompatíveis.
   - Cria `.venv` (app + Word) e, opcionalmente, `.venv-docling` e `.venv-marker`.
     São separados de propósito: Docling e Marker fixam versões diferentes de dependências.
   - Docling/Marker baixam PyTorch e modelos (vários GB; os modelos são baixados na 1ª conversão).
3. Selecione o interpretador `.venv` (`Ctrl+Shift+P > Python: Select Interpreter`).
4. Execute com **F5** (ou `iniciar.bat`). O navegador abre em http://127.0.0.1:5057.

## Uso

- Arraste os PDFs (ou clique na área), escolha o motor e as opções; a fila roda um arquivo por vez.
- Saída: `saida/<nome-do-pdf>/<nome>.md` (+ imagens). Cada item mostra tokens estimados (~caracteres/4).
- Falhou? Botão **Detalhes** mostra a saída técnica e dá para **↻ repetir com outro motor**.
- Erros ficam em `logs/errors.log` (rotativo); botão **Log de erros** abre na tela. `logs/app.log` tem o histórico geral.

## Acelerar Docling/Marker

- **⚡ Modo rápido**: Docling usa `--table-mode fast`; Marker usa `--mode no_ocr` (só texto do PDF); Marker API usa `mode=fast`.
- **🧠 Manter modelos carregados**: um processo do motor (`engine_worker.py`) fica vivo e reaproveita os modelos.
  O 1º arquivo leva o tempo normal; os seguintes ficam muito mais rápidos. Libera a RAM após 20 min sem uso
  (ou no botão **Liberar memória**).
- **Marker API (online)**: roda nos servidores da Datalab (pago, por página). Cole a chave na tela
  (fica em `config.json`, fora do git) ou use a variável `DATALAB_API_KEY`. O PDF sai do seu computador.

## Documentos Word

`.docx`/`.odt` vão direto para o Pandoc (segundos, preserva títulos/listas/tabelas). `.doc`/`.docm`/`.rtf`
passam antes pelo Word para virar `.docx`. O motor selecionado só vale para PDFs.

## Estrutura

| Arquivo | Função |
|---|---|
| `app.py` | Servidor Flask, fila, jobs, log |
| `engines.py` | Docling / Marker (CLI ou worker persistente), Marker API, Word (COM) + Pandoc, documentos Word |
| `engine_worker.py` | Processo que mantém os modelos do Docling/Marker carregados |
| `postprocess.py` | Remove imagens base64, referências, ruído (cabeçalhos/rodapés repetidos, nº de página) |
| `static/index.html` | Interface (modo dia/noite) |
| `static/icon.svg` | Ícone do app |

## Pontos de atenção

- **Flags dos CLIs**: confira com `.venv-docling\Scripts\docling --help` e `.venv-marker\Scripts\marker_single --help`
  se alguma versão mudar opções. Os comandos ficam em `engines.py` (`convert_docling`, `convert_marker`).
- **Word**: fecha instâncias próprias ao terminar, mas PDFs muito complexos podem travar o Word; o tempo limite
  só cobre Docling/Marker/Pandoc. Se travar, finalize `WINWORD.EXE` no Gerenciador de Tarefas.
- **Título/estrutura no Word**: o reflow raramente cria estilos de título. A opção "Detectar títulos" é uma heurística experimental.
- **Limpeza de ruído**: remove linhas idênticas repetidas 4+ vezes (cabeçalhos/rodapés). Desligue se algum texto legítimo sumir.
- O servidor escuta só em `127.0.0.1` (não exposto na rede).

## Versão portátil (pendrive)

1. Numa máquina **com internet**, rode `.\build_portable.bat` (opções: `-SkipMarker`, `-SkipDocling`, `-Dest D:\PDF2MD`).
   Ele baixa Python standalone + Flask/Pandoc + Docling + Marker e já baixa os modelos (conversão de teste).
2. Copie `dist\PDF2MD-Portable` para o pendrive (exFAT/NTFS, USB 3, ~10-15 GB): `robocopy dist\PDF2MD-Portable E:\PDF2MD /E`.
3. Em qualquer Windows 64 bits: `iniciar.bat` no pendrive. Docling/Marker rodam offline; o motor Word só habilita se o PC tiver Word.

Detalhes: caches/modelos ficam em `cache\`; Docling e Marker ficam em `libs\docling` e `libs\marker` (isolados);
`engines.py` detecta a pasta `runtime\` e usa esse Python automaticamente. `PDF2MD_ONLINE=1` permite baixar modelos que faltarem.
