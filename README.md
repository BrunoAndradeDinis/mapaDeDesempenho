# Mapa De Desempenho

Este script Google Apps Script foi desenvolvido em javascript para lidar com a submissão de formulários vinculados a uma planilha do Google que é preenchida através de um formulário pela liderança de nossa equipe. Ele captura os dados submetidos pelo formulário e os adiciona à planilha, além de calcular algumas médias e preencher as respectivas colunas com os resultados.

## Instalação
Abra sua planilha do Google.
No menu superior, clique em "Extensões" > "Apps Script".
Cole o código fornecido neste arquivo na janela do Google Apps Script.
Salve o projeto clicando no ícone de disquete ou pressionando Ctrl + S ou clicando no ícone de disquete.
Feche a janela do Google Apps Script.


## Funcionalidades
```onFormSubmit(e)```
Função que manipula o evento de submissão do formulário.

```e```: Objeto contendo informações sobre o evento de submissão.
```medias(sheet, lastRow)```
Função responsável por calcular e preencher as médias na planilha.

```sheet```: Planilha ativa do Google.
```lastRow```: Última linha preenchida na planilha.
Funções de Cálculo de Médias
Estas funções são responsáveis por calcular e preencher médias específicas na planilha, como média de horário, média de atendimento, média de ambiente de trabalho, média de conhecimento e habilidade, média por colaborador, média da equipe e média ideal.

Cada uma dessas funções recebe como parâmetros a planilha ativa ```(sheet)``` e a última linha preenchida na planilha ```(lastRow)```.

```installTrigger()```
Função para instalar o gatilho de evento de submissão de formulário.

Esta função é usada para instalar um gatilho que aciona a função ```onFormSubmit``` sempre que um formulário vinculado à planilha é submetido.

## Uso

Após instalar o script, ele será acionado automaticamente sempre que um formulário vinculado à planilha for submetido. Os dados submetidos serão adicionados à próxima linha vazia na planilha, que também será criada junto com o envio do formulário, e algumas médias serão calculadas e preenchidas em colunas específicas.
- Média de Horário	
- Média Atendimento	
- Média Ambiente de Trabalho	
- Média Conhecimento/Habilidade	
- Média por Colaborador	
- Média Equipe	
- Média Ideal

## Nota
Devido a ser utilizada no ambito empresarial, optei por não deixar compartilhado o acesso a planilha.
