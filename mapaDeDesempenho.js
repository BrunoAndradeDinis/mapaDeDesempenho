function onFormSubmit(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var formResponses = e.values;

    var lastRow = sheet.getLastRow();

    for (var i = 0; i < formResponses.length; i++) {
        sheet.getRange(lastRow, i + 1).setValue(formResponses[i]);
    }

    medias(sheet, lastRow);
}

function medias(sheet, lastRow) {
    mediaHorario(sheet, lastRow);
    mediAtendimento(sheet, lastRow);
    mediaAmbienteDeTrabalho(sheet, lastRow);
    mediaConhecimentoEHabilidade(sheet, lastRow);
    mediaPorColaborador(sheet, lastRow);
    mediaEquipe(sheet);
    mediaIdeal(sheet);
}

function mediaHorario(sheet, lastRow) {
    var dataRange = sheet.getRange(2, 5, lastRow - 1, 11);
    var data = dataRange.getValues();

    for (var row = 0; row < data.length; row++) {
        var sum = 0;
        var count = 0;
        for (var col = 0; col < data[row].length; col++) {
            var value = parseFloat(data[row][col]);
            if (!isNaN(value)) {
                sum += value;
                count++;
            }
        }
        var average = count > 0 ? sum / count : 0;
        sheet.getRange(row + 2, 16).setValue(average);
    }
}

function mediAtendimento(sheet, lastRow) {
    var rangeG = sheet.getRange("G2:G" + lastRow);
    var rangeH = sheet.getRange("H2:H" + lastRow);
    var rangeI = sheet.getRange("I2:I" + lastRow);

    var valoresG = rangeG.getValues();
    var valoresH = rangeH.getValues();
    var valoresI = rangeI.getValues();

    var rangeQ = sheet.getRange("Q2:Q" + lastRow);

    for (var i = 0; i < lastRow - 1; i++) {
        var valorG = parseFloat(valoresG[i][0]);
        var valorH = parseFloat(valoresH[i][0]);
        var valorI = parseFloat(valoresI[i][0]);
        var media = (valorG + valorH + valorI) / 3;
        rangeQ.getCell(i + 1, 1).setValue(media);
    }
}

function mediaAmbienteDeTrabalho(sheet, lastRow) {
    var rangeJ = sheet.getRange("J2:J" + lastRow);
    var rangeK = sheet.getRange("K2:K" + lastRow);
    var rangeL = sheet.getRange("L2:L" + lastRow);

    var valoresJ = rangeJ.getValues();
    var valoresK = rangeK.getValues();
    var valoresL = rangeL.getValues();

    var rangeR = sheet.getRange("R2:R" + lastRow);

    for (var i = 0; i < lastRow - 1; i++) {
        var valorJ = parseFloat(valoresJ[i][0]);
        var valorK = parseFloat(valoresK[i][0]);
        var valorL = parseFloat(valoresL[i][0]);
        var media = (valorJ + valorK + valorL) / 3;
        rangeR.getCell(i + 1, 1).setValue(media);
    }
}

function mediaConhecimentoEHabilidade(sheet, lastRow) {
    var rangeM = sheet.getRange("M2:M" + lastRow);
    var rangeN = sheet.getRange("N2:N" + lastRow);
    var rangeO = sheet.getRange("O2:O" + lastRow);

    var valoresM = rangeM.getValues();
    var valoresN = rangeN.getValues();
    var valoresO = rangeO.getValues();

    var rangeS = sheet.getRange("S2:S" + lastRow);

    for (var i = 0; i < lastRow - 1; i++) {
        var valorM = parseFloat(valoresM[i][0]);
        var valorN = parseFloat(valoresN[i][0]);
        var valorO = parseFloat(valoresO[i][0]);
        var media = (valorM + valorN + valorO) / 3;
        rangeS.getCell(i + 1, 1).setValue(media);
    }
}

function mediaPorColaborador(sheet, lastRow) {
    var rangeP = sheet.getRange("P2:P" + lastRow);
    var rangeQ = sheet.getRange("Q2:Q" + lastRow);
    var rangeR = sheet.getRange("R2:R" + lastRow);
    var rangeS = sheet.getRange("S2:S" + lastRow);

    var valoresP = rangeP.getValues();
    var valoresQ = rangeQ.getValues();
    var valoresR = rangeR.getValues();
    var valoresS = rangeS.getValues();

    var rangeT = sheet.getRange("T2:T" + lastRow);

    for (var i = 0; i < lastRow - 1; i++) {
        var valorP = parseFloat(valoresP[i][0]);
        var valorQ = parseFloat(valoresQ[i][0]);
        var valorR = parseFloat(valoresR[i][0]);
        var valorS = parseFloat(valoresS[i][0]);
        var media = (valorP + valorQ + valorR + valorS) / 4;
        rangeT.getCell(i + 1, 1).setValue(media);
    }
}

function mediaEquipe(sheet) {
    var rangeT = sheet.getRange("T2:T");
    var valoresT = rangeT.getValues();
    var sum = 0;
    var count = 0;

    for (var i = 0; i < valoresT.length; i++) {
        var value = parseFloat(valoresT[i][0]);
        if (!isNaN(value)) {
            sum += value;
            count++;
        }
    }

    var media = count > 0 ? sum / count : 0;
    var rangeU = sheet.getRange("U2:U" + (count + 1));

    for (var i = 0; i < count; i++) {
        rangeU.getCell(i + 1, 1).setValue(media);
    }
}

function mediaIdeal(sheet) {
    var rangeP = sheet.getRange("P2:P");
    var rangeQ = sheet.getRange("Q2:Q");
    var rangeR = sheet.getRange("R2:R");
    var rangeS = sheet.getRange("S2:S");

    var valoresP = rangeP.getValues();
    var valoresQ = rangeQ.getValues();
    var valoresR = rangeR.getValues();
    var valoresS = rangeS.getValues();

    var rangeV = sheet.getRange("V2:V");

    for (var i = 0; i < rangeV.getNumRows(); i++) {
        var valorP = Math.max(1, parseFloat(valoresP[i][0]));
        var valorQ = Math.max(1, parseFloat(valoresQ[i][0]));
        var valorR = Math.max(1, parseFloat(valoresR[i][0]));
        var valorS = Math.max(1, parseFloat(valoresS[i][0]));
        var media = (valorP + valorQ + valorR + valorS) / 4;
        rangeV.getCell(i + 1, 1).setValue(media);
    }
}

function installTrigger() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onFormSubmit')
        .forSpreadsheet(sheet)
        .onFormSubmit()
        .create();

}


/*
    Nome das colunas
    -Carimbo de data/hora	
    - Período	
    - Colaborador	
    - Fila de Atendimento	
    - Pontualidade	
    - Responsabilidade com o Ponto	
    - Se coloca no lugar da Loja?	
    - Trabalha duro?	
    - Responsabilidade com as Filas	
    - Proatividade
    - Trabalho em Equipe	
    - Bom Relacionamento	
    - Sobe a barra?	
    - Está sempre aprendendo?	
    - Olho no olho	
    - Média de Horário	
    - Média Atendimento	
    - Média Ambiente de Trabalho	
    - Média Conhecimento/Habilidade	
    - Média por Colaborador	
    - Média Equipe	
    - Média Ideal
*/


