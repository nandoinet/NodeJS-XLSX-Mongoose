# NodeJS-XLSX-Mongoose
NodeJS, XLSX e Mongoose X problem Async/Await


Olá! Estou começando a estudar NodeJs e ja estou conseguindo fazer algumas coisasinhas. Mas nos meus experiemntos, estou tentando migrar alguns dados de uma planilha XLSX para o mongo com mongoose.

Eu ja consegui fazer a leitura da planilha e fazer a inclusão de todos os registros no banco, mas como tem alguns CPFs diplicados na planilha, não estou conseuindo fazer uma verificação prévia para ver se o CPF já foi cadastrado ou não...


Ainda não sei bem como fazer o Async/Await funcionar corretamente e por isso não consegi ajusta-los para seguir a sequancia correta.

Da forma como esta o código, o log da execução do código fica assim...

![image](https://user-images.githubusercontent.com/52167139/178641646-fbc67cdb-209b-4f4d-b9f7-88919ae1b70c.png)

E se eu tivesse conseguido ajustar da forma correta, segundo a sequencia dos dados da planilha

![image](https://user-images.githubusercontent.com/52167139/178643335-dd8d093c-79af-4c61-83ff-70338459d8ac.png)

deveria ficar assim... 

![image](https://user-images.githubusercontent.com/52167139/178641703-353a6275-9d0f-4d4b-9d20-6aa1e5605b37.png)

Fico grato pela ajuda...

