export class Montecarlo {

    constructor(distributions) {
        this.distributions = distributions;
        this.variables = new Map();
        this.outputs = new Map();
    }

    isComponentDefined(component) {
        const key = {name: component.name, address: component.address};
        return this.outputs.has(key) || this.variables.has(key);
    }


    addVariable(randomVariable) {
        if (this.isComponentDefined(randomVariable)) {
            return false;
        }
        this.variables.set({name: randomVariable.name, address: randomVariable.address}, randomVariable);
        return true;
    }


    addOutput(output) {
        if (this.isComponentDefined(output)) {
            return false;
        }
        this.outputs.set({name: output.name, address: output.address}, output);
        return true;
    }


    async performSimulation(replicas) {
        const results = new Map();
        const outputs = new Map();
        return Excel.run(async context => {
            const workbook = context.workbook;
            const sheet = workbook.worksheets.getActiveWorksheet();
            for (let i = 0; i < replicas; i++) {

                const variables =
                    Array.from(this.variables)
                        .map(variable => {
                            const address = variable[0].address;
                            return {
                                address: address,
                                value: this.calculateRandomVariableValue(variable[1]),
                                proxy: this.loadCellValue(sheet, address)
                            };
                        });

                await context.sync();

                variables.forEach(({proxy, value}) => proxy.values = [[value]]);

                await context.sync();

                this.outputs.forEach((value, address) => {
                    const cellValue = this.loadCellValue(sheet, value.address);
                    const arrayValues = outputs.get(address);
                    arrayValues
                        ? arrayValues.push(cellValue)
                        : outputs.set(address, [cellValue]);
                });
            }
            await context.sync();

            outputs.forEach((values, address) =>
                results.set(address, values.flatMap(value => value.values[0])));

            // Office.context.ui.displayDialogAsync('https://localhost:3000/dialog.html',{height: 30, width: 30,displayInIframe: true},
            //     function (asyncResult) {
            //        console.log(asyncResult)
            //     });

            await context.sync();

            return Promise.resolve(results);
        });

    }

    loadCellValue(sheet, address) {
        return sheet.getRange(address).load("values");
    }

    calculateRandomVariableValue(variable) {
        let parameters = variable.distributionModel.name === "DISCRETE"
            ? [variable.distributionModel.parameters]
            : [...variable.distributionModel.parameters.values()];
        return this.distributions[variable.distributionModel.name].pdf.apply(this, parameters);
    }

    performStep() {
        this.performSimulation(1);
    }


    clearModel() {
        this.variables.clear();
        this.outputs.clear();
    }
}
