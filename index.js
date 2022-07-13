const XLSX = require('xlsx');
const mongoose = require('mongoose')

// Configurando o Mongoose
mongoose.Promise = global.Promise;
mongoose.connect("mongodb://localhost:27017/dados", {
    useNewUrlParser: true, useUnifiedTopology: true
}).then(() => {
    console.log("MongoDB Conectado...")
}).catch((err) => {
    console.log("Houve um erro ao se conectar ao MongoDB: " + err)
})

const Usuario = mongoose.model('usuario', {
    cpf: { type: String, required: true },
    nome: { type: String, required: true },
    endereco: String,
    numero: Number,
    uf: String
})

console.log("EXPORTAÇÃO DE DADOS - XLSX TO MONGODB")
show_date(new Date(), "INICIADO:   ") // Registra o Inicio do processo

// Endereço do arquivo .XLXS e o indice da pasta de trabalho a ser lida
LerExcel('./dados.xlsx', 0)

async function LerExcel(file, SheetName) {

    const workbook = XLSX.readFile(file);
    //  const workbookSheets = workbook.SheetNames;
    //  console.log(workbookSheets) // Lista as pastas de trabalho da planilha
    const worksheet = workbook.Sheets[workbook.SheetNames[SheetName]]; // 0 é a primeira pasta de trabalho
    const dadosExcel = XLSX.utils.sheet_to_json(worksheet);

    // Lista todos os dados convertidos para JSON
    // console.log(dadosExcel);

    // Percorre todas as colunas para listar os dados das celulas
    for (const cel of dadosExcel) {

        // Valia e armazena os dados
        const usuario = {
            cpf: ckreg(cel['cpf']),
            nome: ckreg(cel['nome']),
            endereco: ckreg(cel['endereco']),
            numero: ckreg(cel['numero']),
            uf: ckreg(cel['uf'])
        }

        await check_user().catch(error => console.log(error.stack));

        async function check_user() {
            await Usuario.exists({ cpf: usuario.cpf }, function (err, doc) {
                if (err) {
                    console.log(err)
                } else {
                    // console.log("Result :", doc) // false
                    if (doc != null) {
                        console.log(usuario.cpf + " registrado")
                    } else {
                        console.log(usuario.cpf + " <-- registrar <--")
                        cad_user()
                    }
                }
            });
        }

        async function cad_user() {
            // SALVA OS DADOS DO USUÁRIO
            try {
                const save_usuario = await new Usuario(usuario).save()
                // res.json({ error: false, usuario: save_usuario })
                console.log(save_usuario.id + ' - ' + save_usuario.nome)
            } catch (err) {
                console.log('ERRO USUÁRIO: ' + err)
            }
        }

    }
            // Registra o fim do processo
            show_date(new Date(), "FINALIZADO: ")

}

// Validação dos dados da planilha
function ckreg(reg) {
    if (reg != null && reg != undefined && reg != "") {
        return reg
    } else {
        return ""
    }
}

// log da data com milisegundos
function show_date(date, msg) {
    options = { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', fractionalSecondDigits: '3' };
    console.log(msg + new Intl.DateTimeFormat("pt-BR", options).format(date));
}