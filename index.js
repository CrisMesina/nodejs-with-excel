import XlsxPopulate from "xlsx-populate";
import readline from "readline";
import fs from "fs";


const usuario = fs.existsSync('./usuarios.xlsx'); // fs.existsSync verifica si un archivo existe en la ruta dada ( ME RETORNA TRUE O FALSE  :O )

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

let existeLibro = usuario;

async function index(){


    console.clear();
    console.log('Que quieres realizar?');
    rl.question(`
        1. ${existeLibro ? 'Ingresar Usuarios' : 'Crear Libro'} \n
        2. Mostrar Listado \n
        3. Modificar datos de usuarios \n
        4. Eliminar datos de usuarios \n
        5. Salir\n`, (opcion) =>{
        switch(opcion){
            case '1':
                if(!existeLibro){
                    console.log('Creando libro...');
                    crearLibro().then(()=>{
                        console.clear();
                        console.log('Libro creado de forma exitosa');
                        setTimeout(()=>{
                            index();

                        }, 1000)
                    })
                }else {
                    console.log('Direccionando a ingreso de usuarios...');
                    setTimeout(() =>{
                        console.clear();
                        ingresarDatos();
                    }, 1000);
                }
                break;
            case '2':
                console.log('Cargando listado...')
                setTimeout(() =>{
                    console.clear();
                    mostrarUsuarios();
                }, 1000);
                break;
            case '3':
                console.log('Opcion en desarrollo');
                break;
            case '4':
                console.log('Opcion en desarrollo');
                break;
            case '5':
                console.log('Saliendo...');
                setTimeout(() =>{
                    console.clear();
                    
                }, 500)
                rl.close();
                break;
            default:
                console.log('Opcion invalida');
                setTimeout(() =>{
                    index();
                }, 1000)
                break;
        }
    })
}

index();


async function crearLibro(){
    const workbook = await XlsxPopulate.fromBlankAsync();

    workbook.sheet(0).cell('A1').value('ID');
    workbook.sheet(0).cell('B1').value('Nombre');
    workbook.sheet(0).cell('C1').value('Apellido');
    workbook.sheet(0).cell('D1').value('Email');

    workbook.toFileAsync('./usuarios.xlsx').then(()=>{
        console.log('Libro creado');
        setTimeout(()=>{
            index();
        }, 500);
    })

    existeLibro = true;

}


async function ingresarDatos() {


    const workbook = await XlsxPopulate.fromFileAsync('./usuarios.xlsx');

    const total = workbook.sheet(0).usedRange().value().length - 1;

    console.clear();
    console.log('Esta es una base datos que almacena información de usuarios.');
    console.log();
    
    rl.question('¿Cuantos usuarios quieres ingresar?: ', (id) =>{
        let contador = 1;
        function agregarUsuario(){
            if(contador <= parseInt(id)){
                rl.question(`Ingrese el nombre del usuario ${total + contador}: `, (nombre) => {
                    rl.question(`Ingrese el apellido del usuario ${total + contador}: `, (apellido) => {
                        workbook.sheet(0).cell(`A${total + contador + 1}`).value(total + contador);
                        workbook.sheet(0).cell(`B${total + contador + 1}`).value(nombre);
                        workbook.sheet(0).cell(`C${total + contador + 1}`).value(apellido);
                        workbook.sheet(0).cell(`D${total + contador + 1}`).formula(`CONCATENATE(B${total + contador + 1},".",C${total + contador + 1},"@correo.com")`);
                        console.clear();
                        contador++;
                        agregarUsuario();
                    });
                });
            } else {
                workbook.toFileAsync('./usuarios.xlsx').then(() => {
                    console.log('Datos guardados en usuarios.xlsx');
                    index();
                });
            }
        }
        agregarUsuario();
    })


}

async function mostrarUsuarios(){
    const workbook = await XlsxPopulate.fromFileAsync('./usuarios.xlsx');
    const datos = workbook.sheet(0).usedRange().value();
    console.table(datos);
    rl.question('Quieres volver? (Si o S)\n', (r) =>{
        if(r.toLowerCase() == 'si' || r.toLowerCase() == 's'){
            console.clear();
            index();
        }
    })
}
