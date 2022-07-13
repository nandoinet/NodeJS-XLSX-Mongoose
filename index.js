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
LerExcel('./dados.xlsx', 0) // 0 é o indice da primeira pasta de trabalho do Excel

async function LerExcel(file, SheetName) {
    const workbook = XLSX.readFile(file);
    //  const workbookSheets = workbook.SheetNames;
    //  console.log(workbookSheets) // Lista as pastas de trabalho da planilha
    const worksheet = workbook.Sheets[workbook.SheetNames[SheetName]]; 
    const dadosExcel = XLSX.utils.sheet_to_json(worksheet);

    // Lista todos os dados convertidos em JSON
    // console.log(dadosExcel);

    // Percorre todas as colunas para listar os dados das celulas
    for (const cel of dadosExcel) {
        const usuario = {
            cpf: ckreg(cel['cpf']),
            nome: ckreg(cel['nome']),
            endereco: ckreg(cel['endereco']),
            numero: ckreg(cel['numero']),
            uf: ckreg(cel['uf']),
        };

        const user_exists = await Usuario.exists({ cpf: usuario.cpf });

        if (!user_exists) {
            const save_usuario = await new Usuario(usuario).save();
            console.log(usuario.cpf + ' <-- registrar <--');
            console.log(save_usuario.id + ' - ' + save_usuario.nome);
        } else {
            console.log(usuario.cpf + ' registrado');
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

// Log data com milisegundos
function show_date(date, msg) {
    options = { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', fractionalSecondDigits: '3' };
    console.log(msg + new Intl.DateTimeFormat("pt-BR", options).format(date));
}
