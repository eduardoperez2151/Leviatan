export class RandomVariable {

    constructor(name, address, distributionModel) {
        this.name = name;
        this.address = address;
        this.distributionModel = distributionModel;
    }


    getInfo() {
        let parameters = ""
        this.distributionModel.parameters
            .forEach((value, key) => parameters += `${key}: ${value}\n`);

        return `Random Variable Name: ${this.name}
            Distribution: ${this.distributionModel.name}
             Parameters ${parameters}`;
    }

}