# calendarioExcel
Calendário por shapes do Excel

Segue uma macro para criar um calendário por shapes no excel.
 
Algo assim:
 
![https://ferramentasexcelvba.files.wordpress.com/2017/11/image0011.png?w=656]

Instruções:
 
1 - Preencher os títulos do eixo vertical
(número da linha sequencial e o label)
 
 
 
2 – Preencher os títulos do eixo horizontal
(Label, posição de início e duração)
 
O label pode ser qualquer string.
Início e duração tem que ser um valor de 0 a 1: é como se fosse uma porcentagem do comprimento total.
 
 
 
 
 
3  - Preencher os dados a serem plotados  e rodar a macro
 
Label (string)
Linha onde estará o retângulo
Início (um percentual do comprimento total)
Duracao (um percentual do comprimento total)
E a cor da caixinha
 
 
 
Para saber a cor, usar a função numcor para descobrir o número.
 
 
Por dentro da caixa preta:
 
O segredo é apenas o comando AddShape, que pode adicionar qualquer dos shapes (retângulo, circunferência, triângulo etc), na posição X, Y, comprimento e altura do quadrado.
 
ActiveSheet.Shapes.AddShape(msoShapeRectangle, posX, posY, Comp, Alt).Select
 
Onde colocar os shapes e como fazer as contas é só matemática simples.
 
Uso técnica semelhante nos trabalhos de Trim e Retrim das máquinas de papel e planejamento operacional florestal.
 
 
