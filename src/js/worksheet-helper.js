const jStat = require('jstat').jStat;

import * as OfficeHelpers from '@microsoft/office-js-helpers';



/*
Crea una nueva planilla por salida del modelo
Calcula las estadisticas
Genera las gráficas
 */
export function createSheet(name, data) {
    const dataLength = data.length;
    Excel.run(async context => {

        //se crea la nueva planilla
        let sheet = context
            .workbook
            .worksheets
            .add(name)
            .load(name);
        await context.sync()
            .catch(() => {
                /*
                si la planilla ya existe da un error por lo que en vez de crearla, reutilizamos la ya existente
                y limpiamos su contenido.
                */
                sheet = context
                    .workbook
                    .worksheets
                    .getItem(name)
                    .load(name);
                sheet.getRange().clear();
                sheet.charts.getItemAt(0).delete();
            });
        // se crean los rangos para los cuales vamos a insertar datos y sus cabezales.
        const replicasHeaders = sheet.getRange("A1:B1").load("values");
        const statisticsHeadersRange = sheet.getRange("D1:D8").load("values");
        const statisticsRangeValues = sheet.getRange("E1:E8").load("values");
        const dataIndexRange = sheet.getRange(`A2:B${dataLength + 1}`).load("values");
        const binsQuantity = Math.ceil(Math.sqrt(dataLength));
        const dataRange = sheet.getRange(`B2:B${dataLength + 1}`);
        setRangeFormat(replicasHeaders, "#008DDD", "#FFFFFF");
        setRangeFormat(statisticsHeadersRange, "#008DDD", "#FFFFFF");
        await context.sync();

        replicasHeaders.values = [[
            'Replica',
            'Output']];

        statisticsHeadersRange.values = [
            ['Mean'],
            ['Median'],
            ['Mode'],
            ['Min'],
            ['Max'],
            ['Range'],
            ['variance'],
            ['Standard Deviation']];

        const jStatMode = jStat.mode(data);
        const mode = (typeof jStatMode === 'object')
            ? jStatMode.join()
            : jStatMode;
        statisticsRangeValues.values = [
            [jStat.mean(data)],
            [jStat.median(data)],
            [mode],
            [jStat.min(data)],
            [jStat.max(data)],
            [jStat.range(data)],
            [jStat.variance(data)],
            [jStat.stdev(data)]];

        dataIndexRange.values = data.map((d, index) => [index + 1, d]);
        //Creamos la Grafica y establecemos sus dimensiones.
        const chart = sheet.charts.add("Histogram", dataRange, "Auto");
        chart.top = 100;
        chart.left = 300;
        chart.height = 400;
        chart.width = 600;
        const series = chart.series
            .getItemAt(0)
            .load(["binOptions/count", "binOptions/type"]);
        await context.sync();

        //agregamos los bins las histograma
        series.binOptions
            .set({type: "BinCount", count: binsQuantity});
        await context.sync();
    })
        .catch(error => {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        });

}


function setRangeFormat(cell, fillColor, fontColor) {
    const cellFormat = cell.format;
    cellFormat.fill.color = fillColor;
    const cellFormatFont = cellFormat.font;
    cellFormatFont.bold = true;
    cellFormatFont.color = fontColor;
}