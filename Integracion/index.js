import fetch from "node-fetch";
import exportFromJSON from "export-from-json";
import Excel from "exceljs";
import * as path from 'path';

const LogIn = async () => {
    const result = await fetch('https://api.talana.com/login-balancer/login/', {
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        },
        method: "POST",
        body: JSON.stringify({
            username: 'integracion-fichas-marcas@cic.cl',
            password: 'SummerEsplendido2969#'
        })
    })
        .then(response => response.json())

    return result.talana_token;
}

/*const GetPeople = async (Token) => {
    const result = await fetch('https://talana.com/es/api/persona', {
        headers: {
            'content-type': 'application/x-www-form-urlencoded',
            'Authorization': Token
        },
    })

    const aux = []
    const data = await result.json();

    data.map((person) => aux.push({
        id: person.id,
        nombre: person.nombre + " " + person.apellidoPaterno + " " + person.apellidoMaterno,
        rut: person.rut
    }))

    return aux;
}*/

const GetMarks = async (Token) => {
    let aditionalPages = true;
    let aux = [];
    const marks = [];
    let currentPage = 1;

    while (aditionalPages) {
        const result = await fetch(`https://talana.com/es/api/mark/?desde=2022-03-01&page=${currentPage}`, {
            headers: {
                'content-type': 'application/x-www-form-urlencoded',
                'Authorization': Token
            },
        })

        const data = await result.json()

        if (data.next) {
            console.log(data['next']);
            aux = aux.concat(data['results']);
        } else {

            aux = aux.concat(data['results']);
            aditionalPages = false;
        }
        currentPage++;
    }

    aux.map((mark) => {
        const fechaHora = mark.TS;
        const aux = fechaHora.split('T')
        const aux2 = aux[1].slice(0, 8);
        marks.push({
            id: mark.person.id,
            name: mark.person.nombre + " " + mark.person.apellidoPaterno + " " + mark.person.apellidoMaterno,
            rut: mark.person.rut,
            date: aux[0],
            time: aux2,
            direction: mark.direction
        })
    });

    const sortedMarks = marks.sort((a, b) => (a.rut < b.rut) ? -1 : (a.rut > b.rut) ? 1 : (a.date < b.date) ? -1 : (a.date > b.date) ? -1 : (a.time < b.time) ? -1 : (a.time > b.time) ? -1 : 0);

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Marks');

    const marksColumns = [
        { key: 'name', header: 'Nombre' },
        { key: 'rut', header: 'Rut' },
        { key: 'date', header: 'Fecha' },
        { key: 'time', header: 'Hora' },
        { key: 'direction', header: 'Entrada(E)/Salida(X)'}
    ];

    worksheet.columns = marksColumns;

    sortedMarks.forEach((mark) => {
        worksheet.addRow(mark);
    });

    const exportPath = path.resolve('C:\\Users\\mikag\\Documents\\CIC\\Integracion', 'marks.xlsx');

    await workbook.xlsx.writeFile(exportPath);
}

const Token = `Token ${await LogIn()}`;
//const People = await GetPeople(Token);
GetMarks(Token);