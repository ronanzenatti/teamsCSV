const fs = require("fs").promises;
const { parse } = require("csv-parse");

const { readdir, stat } = require('fs').promises;
const { sep } = require('path');

const chamadas = [];
let turma = {};
turma.chamadas = [];

let nomeAnterior = '';

function montaChamada(data, nomeTurma) {
    const i = chamadas.indexOf(nomeTurma);

    if (i > -1) {
        chamadas[i].nome.chamadas.push(data);
    } else {
        turma = {};
        turma.chamadas = [];
        turma.nome = nomeTurma;

        chamadas.push(turma);
    }
    console.log(chamadas);
}

function parseArquivo(arquivo, nomeTurma, nomeArquivo, ultimo) {

    var parser = parse(arquivo,
        {
            delimiter: ["\t"],
            encoding: 'utf16le',
            skip_empty_lines: true,
            relaxColumnCount: true,
            relaxQuotes: true,
        }, function (err, records) {
            let alunos = false;
            let dia = true

            const data = {};
            const participantes = [];

            let type = 0;

            records.map(item => {
                const valor = item[0].split('\t');

                if (valor[0].startsWith('﻿Nom') || valor[0].startsWith('Ful') || valor[0].startsWith('Nom')) {
                    alunos = true;
                    if (valor[0].startsWith('﻿Nom'))
                        type = 1;
                    else if (valor[0].startsWith('Ful'))
                        type = 2;
                    else
                        type = 3;

                } else if (alunos) {

                    if (!valor[0]) {
                        alunos = false;
                    }

                    if (valor[0] && !participantes.includes(valor[0])) {
                        participantes.push(valor[0]);
                    }

                    if (dia) {
                        dia = false;
                        if (type == 1) {
                            data.dia = item[2];
                        }
                        else if (type == 3) {
                            data.dia = item[2];
                        }
                        else {
                            data.dia = item[1]
                        }

                    }

                }
            });

            //console.log(nomeArquivo, participantes.length)

            data.participantes = participantes;
            data.arquivo = nomeArquivo;

            turma.nome = nomeTurma;
            turma.chamadas.push(data);

            if (nomeAnterior == nomeTurma || nomeAnterior == '') {

            } else {
                chamadas.push(turma)
                turma = {};
                turma.chamadas = [];
            }
            nomeAnterior = nomeTurma;

            if (ultimo)
                console.log(chamadas)
        });
    // console.log(chamadas)
}

async function lerArquivo() {
    const arquivos = await buscarArquivos();
    ultimo = false;
    arquivos.forEach(async (arq, i) => {
        // Inicio do fim kkk
        if (i + 1 == arquivos.length)
            ultimo = true

        const nomeTurma = arq.split('\\')[1];
        const dadosArquivo = await fs.readFile(arq);
        parseArquivo(dadosArquivo, nomeTurma, arq.split('\\')[2], ultimo)


    })

}

async function listarArquivosDoDiretorio(diretorio, arquivos) {

    if (!arquivos)
        arquivos = [];

    let listaDeArquivos = await readdir(diretorio);
    for (let k in listaDeArquivos) {

        let stat1 = await stat(`${diretorio}${sep}${listaDeArquivos[k]}`);
        if (stat1.isDirectory()) {
            await listarArquivosDoDiretorio(`${diretorio}${sep}${listaDeArquivos[k]}`, arquivos);
        } else {
            let nomeArquivo = `${diretorio}${sep}${listaDeArquivos[k]}`;
            arquivos.push(nomeArquivo);
        }
    }

    return arquivos;

}

async function buscarArquivos() {
    let arquivos = await listarArquivosDoDiretorio('/Users/ronan/OneDrive - Centro Paula Souza/Cisco/LISTA DE PRESENÇA_04_2022'); // coloque o caminho do seu diretorio
    return arquivos;
}

lerArquivo();