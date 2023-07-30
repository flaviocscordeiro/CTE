# CTE
Automação em Python para processar lotes de XML referentes a arquivos de CTe (Conhecimento de Transporte Eletrônico) e gerar um XLSX com todos os dados dos mesmos.

Ao processar o script, o usuário terá que interagir com dois prompts do Windows.
No primeiro, ele terá que indicar o diretório onde estão armazenados os XMLs a serem processados (comvém criar um diretório para cada período de XML a ser analisado).
No próximo prompt, o usuário deverá indicar o local onde deseja salvar os arquivos XLSX (serão dois, um com os arquivos XML que não puderam sem processados - normalmente são arquivos que não foram validados pela RFB (Receita Federal do Brasil). 
E o outro, com todas as informações dos XML processados.

