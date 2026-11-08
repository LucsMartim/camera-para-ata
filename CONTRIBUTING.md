# Como contribuir

Abra uma Issue explicando o comportamento observado, o esperado e a versão do
Windows/Python. Envie exemplos fictícios ou anonimizados; nunca envie gravações
de clientes, transcrições reais, credenciais ou logomarcas sem autorização.

Para propor código:

1. Crie uma branch no seu fork.
2. Faça uma alteração pequena e explique seu caso de uso.
3. Execute `python -m unittest discover -s tests -v`.
4. Abra um pull request com o que mudou e como foi verificado.

Não inclua ambientes virtuais, modelos baixados, arquivos de mídia nem relatórios
gerados. O `.gitignore` ajuda, mas revise os arquivos antes de qualquer commit.
