# Simulador de Investimentos de Fundos Imobiliários

## Objetivo
O simulador tem como objetivo mostrar para o investidor o quanto ele vai ter de **patrimônio acumulado** e **dividendos mensais**, de acordo com a **quantia que ele deseja investir mensalmente**, **anos**, **taxa de rendimento mensal do fundo** e do **percentual de rendimento de sua carteira**.  Além disso, apresenta uma outra opção, onde o usuário poderá ver **cenários**, baseados nas informações já preenchidas, mostrando os **rendimentos e dividendos** de acordo com **anos pré-configurados**.

## Ferramenta utilizada
Microsoft Excel

## Passo a passo para a construção da simulação

Para saber o patrimônio acumulado, foi necessário utilizar a fórmula “VF” que retorna um valor baseado em um investimento futuro. No centro da ferramenta, possui as perguntas que o usuário precisa responder nas células cinzas e baseado nessas informações preenchidas, que são aplicadas as fórmulas. 

![image](https://github.com/user-attachments/assets/a18bc9c4-e784-4618-9eb1-63a02975dc49)

## Aplicando a fórmula VF para descobrir o patrimônio acumulado

No canto esquerdo da ferramenta, utilizei um apoio para fazer o uso da fórmula.
![image](https://github.com/user-attachments/assets/c0249965-6220-408a-ac99-d5a1ef4540a3)


Cada uma das células que são utilizadas para preenchimento das informações, foram nomeadas para facilitar o entendimento e a construção da fórmula. Por exemplo:
Na célula J6, que é utilizada para preencher o quanto o investidor deseja investir por mês, na barra de fórmulas nomeei como “investir_mês”.
![image](https://github.com/user-attachments/assets/b517bb7b-01da-408f-88fe-cff810402ce2)

Tendo feito essas alterações, para descobrir o patrimônio acumulado, utilizei a fórmula VF que possui os seguintes parâmetros: Taxa, NPer, pgto. Onde: 
 
Taxa: Taxa de rendimento mensal do fundo que ele deseja investir
![image](https://github.com/user-attachments/assets/1c1eb992-8365-48c6-80f7-4c153f95db48)


Nper: Quantidade em anos que o investidor deseja investir
Observação: No parâmetro da fórmula, a célula em que está a quantidade de anos, tem que ser multiplicada por 12.
![image](https://github.com/user-attachments/assets/44a45252-6c22-43e0-9c79-3a3e3bc99ac0)


Pgto: Quanto o investidor deseja investir por mês
Observação: No parâmetro da fórmula, a célula em que está a quantia para investir por mês, deve ser multiplicado por -1, para o valor final da fórmula seja positivo.

![image](https://github.com/user-attachments/assets/23ca3c41-6d98-4e33-99c0-8243eb4d1115)


Assim ficou a fórmula em sua estrutura completa

![image](https://github.com/user-attachments/assets/80ae2e36-ec2a-4747-93c8-6731e2b95fe4)

![image](https://github.com/user-attachments/assets/c4232adf-cb3f-40d0-b586-0ef4aea0931e)

### Descobrindo os dividendos mensais

Para descobrir os dividendos mensais é ainda mais fácil, pois foi só multiplicar o valor do patrimônio acumulado por quanto rende a carteira do investidor. Assim, o investidor vai saber quanto vai receber de dividendos mensais, de acordo com a quantia que foi investida em uma determinada quantidade de anos e rendimento do fundo.

![image](https://github.com/user-attachments/assets/a888320c-fd44-4a84-95a2-fe516f9db41d)

![image](https://github.com/user-attachments/assets/4f1d6e36-a727-4245-8daf-419a989eb1a5)

Para esconder essas informações onde estão concentradas as fórmulas, basta apenas mudar a cor da fonte para a cor de fundo e com isso quem for utilizar a ferramenta, não terá esse apoio sendo visualizado, tendo como foco total apenas a utilização do simulador.

## Passo a passo para a construção dos cenários

Os cenários, são exemplos que mostram a quantidade de patrimônio acumulado e dividendos mensais, que são baseados nas informações preenchidas na simulação, em anos pré-configurados.

![image](https://github.com/user-attachments/assets/ba3eecc1-f0db-4211-a500-f3abeab63f3c)

### Patrimônio acumulado

Ao lado das informações escrevi nas células a quantidade de anos, de acordo com os anos pré-configurados, para servir como apoio no uso da fórmula.

![image](https://github.com/user-attachments/assets/556d5e47-63b4-456f-a7ab-7b3002141861)

Para que apareça corretamente os valores de patrimônio acumulado de acordo com os anos pré-configurados, reutilizei a fórmula VF, que possui os parâmetros taxa, nper (quantidade de anos de investimento) e pgto (quantia que deseja investir por mês. Aplicando a fórmula VF nos cenários, fica da seguinte forma:

Taxa: A célula que representa a taxa de rendimento do negócio, localizada na planilha de simulação; nper: A célula que representa a quantidade de anos pré-configurados, multiplicado por 12; pgto: A célula que representa a quantia que deseja investir por mês, localizada na planilha de simulação, multiplicado por -1 para que o resultado seja positivo.

A fórmula foi aplicada para todos esses números a esquerda, conforme mostra a imagem abaixo

![image](https://github.com/user-attachments/assets/0e61d47e-5789-4e8c-a35d-99f0f771252c)

E com isso, temos o patrimônio acumulado com os valores devidamente corretos.

Para esconder esses números que serviram como apoio nas fórmulas, basta apenas mudar a cor da fonte para a cor de fundo e com isso quem for utilizar a ferramenta, não terá esse apoio sendo visualizado, tendo como foco total apenas a utilização dos cenários.

### Dividendos mensais

Os dividendos mensais são descobertos de forma fácil, bastando apenas multiplicar as células de cada patrimônio acumulado nos cenários, pelo rendimento da carteira do investidor, que se encontra na planilha de simulação. Como nomeei as células, incluindo a que é referente ao rendimento da carteira, tive ainda mais facilidade na construção da fórmula. Tendo os parâmentros bem organizados, como mostra a imagem a seguir:

![image](https://github.com/user-attachments/assets/d9cc2add-d572-4a80-9c5e-e12d8f5b01c3)




