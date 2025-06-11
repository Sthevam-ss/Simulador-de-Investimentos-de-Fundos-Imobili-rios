# Simulador de Investimentos de Fundos Imobiliários

## Objetivo
O simulador tem como objetivo mostrar para o investidor o quanto ele vai ter de **patrimônio acumulado** e **dividendos mensais**, de acordo com a **quantia que ele deseja investir mensalmente**, **anos**, **taxa de rendimento mensal do fundo** e do **percentual de rendimento de sua carteira**.  Além disso, apresenta uma outra opção, onde o usuário poderá ver **cenários**, baseados nas informações já preenchidas, mostrando os **rendimentos e dividendos** de acordo com **anos pré-configurados**.

## Ferramenta utilizada
Microsoft Excel

## Passo a passo para a construção da simulação

Para saber o **patrimônio acumulado**, foi necessário utilizar a fórmula **“VF”que retorna um valor baseado em um investimento futuro**. No centro da ferramenta, possui as perguntas que o usuário precisa responder nas **células cinzas** e baseado nessas informações preenchidas, que são aplicadas as fórmulas. 

![image](https://github.com/user-attachments/assets/a18bc9c4-e784-4618-9eb1-63a02975dc49)

## Aplicando a fórmula VF para descobrir o patrimônio acumulado

No canto esquerdo da ferramenta, utilizei um **apoio** para fazer o uso da fórmula.

![image](https://github.com/user-attachments/assets/c0249965-6220-408a-ac99-d5a1ef4540a3)


Cada uma das células que são utilizadas para preenchimento das informações, foram **nomeadas** para facilitar o entendimento e a construção da fórmula. Por exemplo:
Na célula **J6**, **que é utilizada para preencher o quanto o investidor deseja investir por mês**, na barra de fórmulas nomeei como **“investir_mês”**.

![image](https://github.com/user-attachments/assets/b517bb7b-01da-408f-88fe-cff810402ce2)

Tendo feito essas alterações, para descobrir o **patrimônio acumulado**, utilizei a fórmula **VF** que possui os seguintes parâmetros: **Taxa, NPer, pgto.** Onde: 
 
**Taxa:** Taxa de rendimento mensal do fundo que ele deseja investir

![image](https://github.com/user-attachments/assets/1c1eb992-8365-48c6-80f7-4c153f95db48)


**Nper:** Quantidade em anos que o investidor deseja investir

**Observação:** No parâmetro da fórmula, a célula em que está a quantidade de anos, tem que ser multiplicada por 12.

![image](https://github.com/user-attachments/assets/44a45252-6c22-43e0-9c79-3a3e3bc99ac0)


**Pgto:** Quanto o investidor deseja investir por mês

**Observação:** No parâmetro da fórmula, a célula em que está a quantia para investir por mês, deve ser multiplicado por -1, para o valor final da fórmula ser positivo.

![image](https://github.com/user-attachments/assets/23ca3c41-6d98-4e33-99c0-8243eb4d1115)


Assim ficou a fórmula em sua estrutura completa

![image](https://github.com/user-attachments/assets/80ae2e36-ec2a-4747-93c8-6731e2b95fe4)

![image](https://github.com/user-attachments/assets/c4232adf-cb3f-40d0-b586-0ef4aea0931e)

### Descobrindo os dividendos mensais

Para descobrir os dividendos mensais é ainda mais fácil, pois foi só **multiplicar o valor do patrimônio acumulado por quanto rende a carteira do investidor.** Assim, o investidor vai saber quanto vai receber de **dividendos mensais**, de acordo com a **quantia que foi investida em uma determinada quantidade de anos e rendimento do fundo.**

![image](https://github.com/user-attachments/assets/a888320c-fd44-4a84-95a2-fe516f9db41d)

![image](https://github.com/user-attachments/assets/4f1d6e36-a727-4245-8daf-419a989eb1a5)

Para esconder essas informações onde estão concentradas as fórmulas, basta apenas **mudar a cor da fonte para a cor de fundo** e com isso quem for utilizar a ferramenta, não terá esse apoio sendo visualizado, tendo como foco total apenas a utilização do simulador.

## Passo a passo para a construção dos cenários

Os cenários, são exemplos que mostram a **quantidade de patrimônio acumulado e dividendos mensais**, que são baseados nas **informações preenchidas na simulação**, de acordo com **anos pré-configurados.**

![image](https://github.com/user-attachments/assets/ba3eecc1-f0db-4211-a500-f3abeab63f3c)

### Patrimônio acumulado

Ao lado das informações escrevi nas células a quantidade de anos, de acordo com os anos pré-configurados, para servir como apoio no uso da fórmula.

![image](https://github.com/user-attachments/assets/556d5e47-63b4-456f-a7ab-7b3002141861)

Para que apareça corretamente os valores de **patrimônio acumulado de acordo com os anos pré-configurados**, reutilizei a fórmula VF, que possui os parâmetros **taxa**, **nper** (quantidade de anos de investimento) e **pgto** (quantia que deseja investir por mês). Aplicando a fórmula VF nos cenários, fica da seguinte forma:

**Taxa:** A célula que representa a taxa de rendimento do negócio, localizada na planilha de simulação

**nper:** A célula que representa a quantidade de anos pré-configurados, multiplicado por 12

**pgto:** A célula que representa a quantia que deseja investir por mês, localizada na planilha de simulação, multiplicado por -1 para que o resultado seja positivo.

A fórmula foi aplicada para todos esses números a esquerda, conforme mostra a imagem abaixo

![image](https://github.com/user-attachments/assets/0e61d47e-5789-4e8c-a35d-99f0f771252c)

E com isso, temos o **patrimônio acumulado** com os valores devidamente corretos.

Para esconder esses números que serviram como apoio nas fórmulas, basta apenas **mudar a cor da fonte para a cor de fundo** e com isso quem for utilizar a ferramenta, não terá esse apoio sendo visualizado, tendo como foco total apenas a utilização dos cenários.

### Dividendos mensais

Os dividendos mensais são descobertos de forma fácil, bastando apenas **multiplicar as células de cada patrimônio acumulado nos cenários, pelo rendimento da carteira do investidor**, que se encontra na **planilha de simulação**. Como **nomeei as células, incluindo a que é referente ao rendimento da carteira**, tive ainda mais facilidade na construção da fórmula. Tendo os parâmentros bem organizados, como mostra a imagem a seguir:

![image](https://github.com/user-attachments/assets/d9cc2add-d572-4a80-9c5e-e12d8f5b01c3)

## Construção do Layout da Ferramenta

**1° Passo:** Alterei toda a cor da planilha para a segunda tonalidade de branco

![image](https://github.com/user-attachments/assets/1a0d8ea1-9cb1-446f-ba1a-2e732245c5fc)


**2° Passo:** Selecionei a área que desejei construir o layout e em seguida, segurando as teclas do teclado CTRL + 1, indo na **aba borda**, selecionei a **opção contorno** e **cliquei em ok**. Foi necessário também **alterar a cor de fundo da área como branco.**

![image](https://github.com/user-attachments/assets/6add4f7f-ef0b-48c6-940f-4a13926ce5f3)

![image](https://github.com/user-attachments/assets/da58c643-8c93-4384-bc6e-44eb19cd77e4)

**3° Passo:** A maior parte dos elementos que utilizei no layout do simulador, foi utilizada a **forma retângulo: canto arredondado**. Que pode ser encontrado facilmente na **guia inserir** e em seguida na **opção formas.** E com isso, apenas **modifiquei as cores e os textos** deles para a construção dos **menus de navegação e título**.

![image](https://github.com/user-attachments/assets/eb2014f9-52da-4d38-969b-da5503ae4d42)

![image](https://github.com/user-attachments/assets/dd922893-ca60-4d77-96ab-99f56c062541)

No meio do Layout, apenas **escrevi em negrito os títulos das informações e usei um cinza mais escuro para os campos que o investidor precisa preencher**. Ao lado de cada campo, apenas utilizei um **retângulo de cantos arredondados com tamanho menor**, para dar destaque. Inclusive, esses retângulos também foram utilizados como base para a **criação da logo do projeto**.

![image](https://github.com/user-attachments/assets/913aaf43-c414-46f3-86ad-448a036362be)

Foram utilizadas **caixas de texto para a construção do nome do simulador e também do título principal do layout**. **A caixa de texto pode ser encontrada facilmente na aba inserir**. Lembrando que **é necessário retirar o contorno e o fundo**, para que fique amostra somente os textos.

![image](https://github.com/user-attachments/assets/fe591e93-04df-4a3f-b62e-53d8260b40c4)

**As informações numéricas* que aparecem nos **retângulos de cantos arredondados**, também são caixas de texto. A diferença, é que elas refereciam os valores obtidos através da **fórmula VF**, em relação ao **patrimônio acumulado**, **dividendos mensais e rendimento da carteira**.
Pra fazer isso, basta apenas **clicar em cima do texto que desejar** e em seguida na barra de fórmulas, usar o **símbolo de igual (=)** e em seguida a **célula que consta a informação que deseja mostrar**. No exemplo abaixo, mostra a célula que possui o **valor do patrimônio acumulado**, neste caso a **célula A23**.

![image](https://github.com/user-attachments/assets/1a934dad-7390-4894-873d-19fc699a1011)

**Os ícones podem ser facilmente encontrados na guia inserir**. E basta apenas pesquisar o tipo de ícone que desejar e **clicar em inserir**. Eles podem ser formatados da forma que quiser, da mesma forma que são formatados as formas retângulares encontradas no layout.

![image](https://github.com/user-attachments/assets/2bc97d88-f533-4483-9ae8-6089602e988f)

Depois de ter inserido todas essas informações, foi necessário **agrupar** para que todas as **formas**, **ícones e caixa de texto**, se tornassem **um objeto só**. Selecionando cada uma das informações com a **tecla CTRL do teclado**, basta apenas **clicar com o botão direito após as seleções e clicar em agrupar**. Foi necessário fazer isso para **todos os outros retângulos de cantos arrendondados separadamente**, para **todos se agruparem e se tornarem apenas um objeto.**

![image](https://github.com/user-attachments/assets/ab653d90-01ca-4899-8a98-1be7dd4afd2f)

## Navegação

Para criar **múltiplas planilhas** e ter como **navegar entre elas através dos botões de menu**, foi necessário apenas **criar uma cópia da planilha atual de simulação e criar hyperlinks** que fornecem ação a esses botões.

**1° Passo:** Para copiar a planilha atual, basta apenas clicar com o botão direito do mouse em cima dela, clicar na opção **"mover ou copiar"** e depois selecionar **"criar uma cópia"** e **clicar em OK.**

![image](https://github.com/user-attachments/assets/01389b89-35c8-4e5f-9a3d-b7c8d8a2d635)
![image](https://github.com/user-attachments/assets/6a40f8b6-62eb-4021-8747-e198b94e170d)

**2° Passo:** Agora, com a planilha copiada, apenas foi apagado as informações da cópia e construído toda a estrutura dos cenários de investimentos, como foi explicado no começo desta documentação.

**3° Passo:** Para que os **botões tenham ações**, é necessário **adicionar hyperlinks**, para que quando um botão for selecionado, ele direcione para planilha que desejar.

Para isso, **na planilha de simulação**, basta **clicar com o botão direito sobre "cenários"**, **clicar em "links"**, **"colocar neste documento"** e **selecionar a planilha de cenários** e em seguida **clicar em OK.**

![image](https://github.com/user-attachments/assets/7dcabc8f-6ada-4b53-807a-5ddb54eb491d)
![image](https://github.com/user-attachments/assets/2f01988f-6659-445c-bda4-dcf57924e076)

Na **planilha de cenários**, foi necessário fazer o mesmo passo a passo, mas aplicando no **botão simulação e vinculando a planilha de simulação.**

![image](https://github.com/user-attachments/assets/b3511653-c596-4591-9274-d1074be57934)
![image](https://github.com/user-attachments/assets/0e1fa1ab-9b2b-475a-a033-be8559f430a9)

Assim, ao clicar nos botões, terá o efeito de menu onde cada botão te levará a planilhas diferentes.

Além dessas planilhas, criei uma planilha como **tela inicial**, efetuando todos os passos anteriores de **hyperlinks**, **caixas de texto e retângulos com cantos arredondados.** **O ícone de casa** apresentado nas demais planilhas, com o hyperlink configurado, levam diretamente a essa **tela inicial.** 

## Tela Inicial
![image](https://github.com/user-attachments/assets/10d5da7d-a4c3-4668-a99d-36da4319400c)

## Tela de Simulação

![image](https://github.com/user-attachments/assets/8df4c9a1-d523-4223-9932-8a013700ce78)

## Tela de Cenários

![image](https://github.com/user-attachments/assets/d53e9484-38a1-4b33-bc73-8f05caf6e6bf)

O último detalhe, é que para a **planilha ficar com um visual de sistema**, foi necessário ir até a **guia exibir** e **desabilitar **as seguintes opções: **Linhas de Grade**, **Barra de Fórmulas** e **Títulos**. Além de utilizar a **combinação de teclas CTRL + F1, para habilitar o modo tela cheia.**

![image](https://github.com/user-attachments/assets/908ed834-256c-424b-a71f-75bea11ca71c)

## Projeto com todas as opções habilitadas

![image](https://github.com/user-attachments/assets/0fa470b5-e380-4a31-a147-2bc27dbb3de2)


## Projeto com todas as opções desabilitadas

![image](https://github.com/user-attachments/assets/c4f99a81-faf0-4c6f-9855-c8889d8fbaf8)

E com todos esses conhecimentos aplicados, o projeto de simulador de fundos imobiliários foi concluído.
































