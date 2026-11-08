# Câmera para ATA

Aplicativo local para Windows que transforma o **áudio de vídeos gravados** em
transcrição e rascunhos de atas de reunião. Possui interface gráfica em Python/Tkinter.

O sistema trabalha com a fala gravada: **não interpreta as imagens**, não identifica
rostos e não acessa câmeras ao vivo. Os vídeos precisam conter áudio compatível.

## Funcionalidades

- Seleção e organização de vários vídeos de uma reunião.
- Extração de áudio por vídeo e criação de um MP3 contínuo.
- Transcrição em português com faster-whisper, usando CPU e int8.
- Referência ao vídeo de origem e à posição de cada trecho.
- Identificação aproximada de horário quando o nome do vídeo contém data e hora.
- Dados da reunião e logomarca opcional, informados pelo usuário.
- Exportação de transcrição, rascunho Markdown e documentos Word.

A seleção dos assuntos e possíveis encaminhamentos usa regras simples de texto.
Não há identificação de falantes nem um modelo de linguagem para resumir a reunião.
As atas precisam de revisão humana antes de aprovação ou entrega.

## Instalação no Windows

Pré-requisito recomendado: **Python 3.12 de 64 bits**, com Tkinter e pip.
Baixe em [python.org](https://www.python.org/downloads/windows/).

1. No GitHub, escolha **Code → Download ZIP** e extraia a pasta.
2. Execute `INSTALAR.bat` e aguarde a instalação das dependências.
3. Execute `INICIAR_SISTEMA.bat` para abrir o aplicativo.

Alternativa pelo terminal, dentro da pasta do projeto:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe sistema_ata.py
```

A primeira transcrição baixa o modelo `medium`, caso ele ainda não esteja disponível
na pasta `.modelos`. É necessário acesso à internet para instalar dependências e
baixar o modelo. O download pode ser grande e a transcrição na CPU pode demorar.
O executável FFmpeg é obtido pela biblioteca imageio-ffmpeg.
Os intervalos de versão estão no requirements.txt; não representam um lock de todas
as dependências nem uma garantia de compatibilidade com futuras versões.

## Como usar

1. Adicione os vídeos da reunião e confira a ordem com **Subir** e **Descer**.
2. Preencha cliente, título, data, horário, local e participantes.
3. Se desejar, selecione uma logomarca própria em PNG ou JPG.
4. Escolha a pasta de saída e clique em **Processar reunião completa**.
5. Confira a transcrição no áudio e revise a ata antes de utilizar.

São aceitos na seleção arquivos MP4, MKV, AVI e MOV, sujeitos aos codecs disponíveis.
Para recuperar o início da gravação, o nome deve conter um padrão como
`2026-01-10-14-30-00`. Sem esse padrão, o relatório indica horário não identificado.
Os horários dependem do nome do arquivo; não são uma verificação do relógio da câmera.

## Arquivos gerados localmente

| Arquivo | Finalidade |
| --- | --- |
| `AUDIO_REUNIAO_COMPLETA.mp3` | Áudio reunido |
| `audios_por_video/` | MP3 separado por vídeo |
| `TRANSCRICAO_COMPLETA.txt` | Falas com referências de origem |
| `ATA_RASCUNHO.md` | Trechos e possíveis encaminhamentos |
| `RELATORIO_TRANSCRICAO.docx` | Transcrição formatada para conferência |
| `ATA_REUNIAO.docx` | Ata preliminar para revisão |
| `dados_processamento.json` | Metadados e segmentos transcritos |

**Esses arquivos não fazem parte do repositório.** O código público não inclui
vídeos, áudios, reuniões, nomes de clientes, logomarcas ou documentos já processados.
Não é necessário baixar ou copiar o ambiente de execução do computador de origem.

## Privacidade

O aplicativo processa as gravações localmente e não implementa envio de vídeos ou
áudios a uma API de transcrição. As bibliotecas e o modelo são baixados de serviços
externos durante a instalação e o primeiro uso. Não se promete funcionamento
totalmente offline antes desses downloads.

Use gravações que você está autorizado a processar. Proteja a pasta de saída:
ela contém falas, dados da reunião e caminhos locais dos vídeos. Não coloque
gravações nem relatórios em pastas de código destinadas à publicação.

## Desenvolvimento e validação

```powershell
py -3.12 -m unittest discover -s tests -v
```

Os testes usam dados fictícios e verificam funções de nomes, datas e duração,
saídas de texto/JSON e tratamento de falha de conversão. Não baixam modelos nem
transcrevem gravações. O teste de uma reunião completa e a validação visual da
interface devem ser feitos localmente antes de uso em produção.

Limitações desta versão: processamento sequencial, sem cancelamento pela interface,
sem identificação de falantes e sem integração ao vivo com câmeras. Um segmento
que atravessa a fronteira entre dois vídeos recebe a referência do vídeo inicial.

## Código aberto e hospedagem

Este repositório disponibiliza o **código-fonte**, para baixar e executar localmente.
O aplicativo desktop não funciona como site no GitHub Pages. Uma versão web exigiria
outro projeto de implantação, com servidor, armazenamento e controle de acesso.

Veja [CONTRIBUTING.md](CONTRIBUTING.md) para contribuir e
[ROADMAP.md](ROADMAP.md) para sugestões de evolução.

## Dependências e licença

Código do projeto sob [licença MIT](LICENSE). Bibliotecas, modelos e FFmpeg mantêm
suas próprias licenças; não são redistribuídos neste repositório.

- [faster-whisper](https://github.com/SYSTRAN/faster-whisper)
- [PyAV](https://github.com/PyAV-Org/PyAV)
- [imageio-ffmpeg](https://github.com/imageio/imageio-ffmpeg)
- [python-docx](https://github.com/python-openxml/python-docx)

Projeto independente, sem vínculo oficial com fabricantes de câmeras.
