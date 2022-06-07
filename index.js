const fs = require("fs").promises;
const { parse } = require("csv-parse");

const { readdir, stat } = require('fs').promises;
const { sep } = require('path');
const xlsx = require("node-xlsx");

const path = '/Users/ronan/OneDrive - Centro Paula Souza/Cisco/LISTA DE PRESENÇA_04_2022';

const chamadas = [];
let turma = {};
turma.chamadas = [];

let arquivos = [];

let nomeAnterior = '';
let ultimo = 0;

let i = 0;


function exportacao(planilha) {
    let exportacao = [];

    planilha.forEach(item => {
        let temp = {}
        temp.name = item.name;
        let data = [];
        item.data.forEach(aluno => {
            data.push(aluno.split(';'));
        })
        temp.data = data;
        exportacao.push(temp)
    });

    console.log(` -> Iniciando a Gravação do arquivo:`);
    var buffer = xlsx.build(exportacao); // Returns a buffer

    fs.writeFile(`CCNA.xlsx`, buffer);
    console.log(` -> Arquivo gravado com sucesso!`);

}

function montarExcel() {
    let planilha = [];
    let turmaTemp = {};
    let nomes = [];
    chamadas.forEach(turma => {
        nomes = [];
        nomes.push(turma.nome)
        turma.chamadas.forEach(dia => {

            dia.participantes.forEach(aluno => {
                if (!nomes.includes(aluno)) {
                    nomes.push(aluno)
                }
            }); // Dia

        }); // Turma

        turmaTemp.name = turma.nome;
        turmaTemp.data = [];
        turma.chamadas.forEach(dia => {
            nomes.forEach((nome, i) => {
                if (i == 0) {
                    nomes[i] += `;${dia.dia}`
                } else {
                    //console.log(dia.participantes.includes(nome))
                    if (dia.participantes.includes(nome.split(';')[0])) {
                        nomes[i] += `;X`
                    } else {
                        nomes[i] += `;-`
                    }
                }
            }); // Turma
            turmaTemp.data = nomes;
        });
        planilha.push(turmaTemp);
        turmaTemp = {};
    }); //Chamada
    // console.log(planilha)
    exportacao(planilha);
}

async function processar() {

    const dadosArquivo = await fs.readFile(arquivos[i]);
    const nomeTurma = arquivos[i].split('\\')[1];
    const nomeArquivo = arquivos[i].split('\\')[2];

    if (nomeAnterior == nomeTurma || nomeAnterior == '') {

    } else {
        turma.nome = nomeAnterior;
        // console.log(turma)
        chamadas.push(turma)
        turma = {};
        turma.chamadas = [];
    }

    var parser = parse(dadosArquivo,
        {
            delimiter: ["\t"],
            encoding: 'utf16le',
            // skip_empty_lines: true,
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
                            data.dia = item[2].split(' ')[0].split(',')[0];
                        }
                        else if (type == 3) {
                            data.dia = item[2].split(' ')[0].split(',')[0];
                        }
                        else {
                            data.dia = item[1].split(' ')[0].split(',')[0];
                        }
                    }
                }
            });

            //depois
            data.participantes = participantes;
            data.arquivo = nomeArquivo;
            turma.chamadas.push(data);

            nomeAnterior = nomeTurma;
            i++;

            if (i < ultimo) {
                processar();
            } else {
                turma.nome = nomeAnterior;
                chamadas.push(turma)
                //  console.log(chamadas);

                montarExcel()
            }

        });

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
    arquivos = await listarArquivosDoDiretorio(path);
    ultimo = arquivos.length;

    processar();

}

buscarArquivos();