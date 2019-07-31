export class OutputVariable {

    constructor(name, address, formula) {
        this.name = name;
        this.address = address;
        this.formula = formula;
    }


    getInfo() {

        return `Output Variable Name: ${this.name}\nOutput Value: ${this.value}`
    }
}