import * as OfficeHelpers from '@microsoft/office-js-helpers';
import * as WorkbookUtils from './worksheet-helper'
import {distributions} from "./distributions";
import {RandomVariable} from "./models/random-variable";
import {DistributionModel} from "./models/distribution-model";
import {OutputVariable} from "./models/ouput-variable";
import {Montecarlo} from "./models/montecarlo";

(() => {

    const clickEvent = "click";
    const displayNoneStyle = "d-none";

    const $menu = $("#menu");
    const $btnBack = $("#btnBack");
    const $mcStep = $("#mcStep");
    const $mcList = $("#mcList");
    const $mcSimulate = $("#mcSimulate");
    const $mcClear = $("#mcClear");
    const $replicateInput = $("#replicateInput");
    const $btnOutput = $("#btnOutput");
    const $outputForm = $("#outputForm");
    const $outputNameInput = $('#outputNameInput');
    const $backContainer = $("#backContainer");
    const $btnMontecarlo = $("#btnMontecarlo");
    const $montecarloForm = $("#montecarloForm");
    const $btnRandomVariable = $("#btnRandomVariable");
    const $randomVariableForm = $("#randomVariableForm");

    const montecarlo = new Montecarlo(distributions);

    /*
    Agrega el evento click al botón($button), para hacer visible el formulario($form).
     */
    function showFormOnClick($button, $form) {
        $button.on(clickEvent, () => {
            $form.removeClass(displayNoneStyle);
            $backContainer.removeClass(displayNoneStyle);
            $menu.addClass(displayNoneStyle);
        });
    }

    /*
    Dada una distribución, muestra la información de un componente (Variable Aletoría o Salida)
    en el formulario de la Simulación del Modelo de Montecarlo.
     */
    function createMontecarloUIComponents(type, distribution, name, colorClass) {
        const distributionInfo = distribution
            ? "<div>" +
            "   <span class='font-weight-bold'>Distribution: </span>" +
            "   <span>" + distribution + "</span>" +
            "</div>"
            : "";
        return "<div class='mc-item-list " + colorClass + "'>\n" +
            "   <div class='mc-item-info'>\n" +
            "       <div>\n" +
            "           <span class='font-weight-bold'>Type: </span>\n" +
            "           <span>" + type + "</span>\n" +
            "       </div>" +
            distributionInfo +
            "       <div>" +
            "           <span class='font-weight-bold'>Name: </span>" +
            "           <span>" + name + "</span>" +
            "       </div>" +
            "   </div>" +
            "</div>";
    }


    /*
    Genera un componente visual para cada una de las Variables aleatorias o Salidas.
    */
    function showMontecarloModel() {
        $mcList.empty();
        montecarlo.variables.forEach(rv => {
            $mcList.append(createMontecarloUIComponents("Random Variable", rv.distributionModel.name, rv.name, "bg-success"));
        });

        montecarlo.outputs.forEach(p => {
            $mcList.append(createMontecarloUIComponents("Output Variable", undefined, p.name, "bg-primary"));
        });
    }

    /*
    Muestra el formulario de Simulación
     */
    function showMontecarloForm() {
        $btnMontecarlo.on(clickEvent, () => {
            $montecarloForm.removeClass(displayNoneStyle);
            $backContainer.removeClass(displayNoneStyle);
            $menu.addClass(displayNoneStyle);
            showMontecarloModel();
        });


        //Al hacer click en el botón "Clear" se limpia el modelo de montecarlo junto con la UI
        $mcClear.on(clickEvent,
            () => {
                montecarlo.clearModel();
                $(".mc-item-list").remove();
                clearWorkbook();
            });

        //Ejecuta un paso de simulación
        $mcStep.on(clickEvent,
            () => montecarlo.performStep());

        //Inicia la simulación de Montecarlo
        $mcSimulate.on(clickEvent,
            () => {
                const replicate = $replicateInput.val();
                montecarlo.performSimulation(replicate)
                    .then(results => {
                        results.forEach((value, key) =>
                            WorkbookUtils.createSheet(key.name, value));
                    })
                    .catch(error => {
                        OfficeHelpers.UI.notify(error);
                        OfficeHelpers.Utilities.log(error);
                    });
            });
    }

    /*
    Al hacer Click en el botón atras se oculta el formulario.
     */
    function hideFormsOnBack() {
        $btnBack.on(clickEvent, () => {
            $menu.removeClass(displayNoneStyle);
            $randomVariableForm.addClass(displayNoneStyle);
            $outputForm.addClass(displayNoneStyle);
            $montecarloForm.addClass(displayNoneStyle);
            $backContainer.addClass(displayNoneStyle);
            $mcList.empty();
        });
    }

    /*
    Agrega el Combo de Distribuciones los nombres de las mismas configuradas en el archivo
    distributions.js
     */
    function setDistributionSelectOptions() {
        const $distributionSelect = $('#distributionSelect');
        Object.entries(distributions)
            .forEach(distribution => {
                const option = new Option(distribution[1].name, distribution[0]);
                $distributionSelect.append(option);
            });
    }

    /*
    Genera el componente de UI para el fomulario de creación de la variable aleatoria, cuando se selecciona
    La distribución discreta.
     */
    function addDiscreteParameters() {

        $('<hr>')
            .insertBefore($(".btn").first());
        $('<div class="parameter form-group">' +
            '<label>Value</label>' +
            '<input  type="number" class="value form-control" required/>' +
            '<div class="invalid-feedback">Value is required.</div>' +
            '</div>')
            .insertBefore($(".btn").first());
        $('<div class="parameter form-group">' +
            '<label>Probability</label>' +
            '<input  type="number" class="probability form-control" min="0" max="1" step="0.0001" value="0" required/>' +
            '<div class="invalid-feedback">Probability is required.</div>' +
            '</div>')
            .insertBefore($(".btn").first());
    }


    /*
    Genera dinámicamente los inputs para los parametros al seleccionar la distribución.
     */
    function createDistributionParameterInputs() {
        const $distributionSelect = $('#distributionSelect');
        $distributionSelect.on('change', () => {
            const selectedDistribution = $distributionSelect.val();
            $('div').remove('#randomVariableForm .parameter');
            const distribution = distributions[selectedDistribution];

            if (distribution.name === "Discrete") {
                addDiscreteParameters();
                const $firstButton = $(".btn").first();
                $('<button id="discretePlus" type="button" class="btn btn-primary m-2">'
                    + '<i class="fa fa-plus"></i></button>')
                    .on("click", () => {
                        addDiscreteParameters();
                    })
                    .insertBefore($firstButton);

                $('<button id="discreteMinus" type="button" class="btn btn-danger m-2">'
                    + '<i class="fa fa-minus"></i></button>')
                    .on("click", () => {
                        const $parameter = $('.parameter');
                        if ($parameter.length > 2) {
                            $('hr').last().remove();
                            $parameter.slice(-2).remove();
                        }
                    })
                    .insertBefore($firstButton);
                return
            }

            $('hr').remove();
            $("#discretePlus").remove();
            $("#discreteMinus").remove();

            distribution
                .parameters
                .forEach(parameter => {
                    const displayName = parameter.displayName;
                    $('<div class="parameter form-group">\n' +
                        '<label for="' + displayName + 'Input">' + displayName + '</label>\n' +
                        '<input  type="number" value="' + parameter.defaultValue + '" class="form-control" id="' + parameter.name + 'Input" required/>\n' +
                        '<div class="invalid-feedback">\n' +
                        displayName + ' is required.\n' +
                        '</div>\n' +
                        '</div>')
                        .insertBefore($(".btn").first());
                });

        });
    }


    /*
    Activa la validación de los formularios utilizando boostrap
     */
    function loadFormValidation() {
        window.addEventListener('load', function () {
            const forms = document.getElementsByClassName('needs-validation');
            Array.prototype.filter.call(forms, function (form) {
                form.addEventListener('submit', function (event) {
                    if (form.checkValidity() === false) {
                        event.preventDefault();
                        event.stopPropagation();
                    }
                    form.classList.add('was-validated');
                }, false);
            });
        }, false);
    }

    /*
    Dado un rango de celdas se seleciona unicamente la primera.
    Se devuelve el proxy de la celda con las propiedades address y values
    listas para ser utilizadas.
     */
    function getFirstSelectedCell(workbook) {
        const selectedRange = workbook.getSelectedRange();
        let cell = selectedRange.getCell(0, 0);
        cell.load(['address', 'values']);
        return cell;
    }

    /*
    Obtiene los valores de los parámetros cuando se selecciona, la distribución discreta.
     */
    function getDiscreteDistributionParameters() {
        const parameters = $(".value, .probability");
        const discreteParameters = new Map()
        for (let index = 0; index < parameters.length; index += 2) {
            const value = $(parameters[index]).val();
            const probability = $(parameters[index + 1]).val();
            discreteParameters.set(value, probability);
        }
        return discreteParameters;
    }

    /*
    Dada la distribución selecionada y los parametros ingresados, se devuelve
    un objeto DistributionModel que representa una distribucion con su nombre y parámetros.
     */
    function getSelectedDistributionModel() {
        let name = $('#distributionSelect').val();
        let distribution = distributions[name];

        let parameters =
            name === "DISCRETE" // Se agrega un tratamiento especial para la distribución completa.
                ? getDiscreteDistributionParameters()
                : new Map(distribution.parameters.map(parameter => [parameter.name, parseInt($("#" + parameter.name + "Input").val())]));
        return new DistributionModel(name, parameters)
    }

    /*
    Dado un rango se establece el color de relleno de las celdas y de fuente.
     */
    function setRangeFormat(cell, fillColor, fontColor) {
        const cellFormat = cell.format;
        cellFormat.fill.color = fillColor;
        const cellFormatFont = cellFormat.font;
        cellFormatFont.bold = true;
        cellFormatFont.color = fontColor;
    }

    /*
    Dado un modelo de distribucion y una celda se crea un objeto RandomVariable que representa una variable
    Aleatoria.
     */
    function createRandomVariable(distributionModel, cell) {
        const name = $('#randomVariableNameInput').val();
        const address = cell.address;
        return new RandomVariable(name, address, distributionModel)
    }

    /*
    Dada una celda crea un objeto que representa una salida del modelo de montecarlo.
     */
    function createOutputVariable(cell) {
        const name = $outputNameInput.val();
        const address = cell.address;
        const value = cell.values;
        return new OutputVariable(name, address, value)
    }

    /*
    Agrega un comentario a la celda especificada.
     */
    function addComments(component, workbook, cell) {
        const parameterCommentInfo = component.getInfo();
        const comments = workbook.comments;
        comments.add(parameterCommentInfo, cell);

    }

    /*
    Agrega la variable aleatoria al Model de Montecarlo
     */
    function performMontecarloRandomVariableAddition() {
        $randomVariableForm.submit(event => {
            event.preventDefault();
            const form = document.getElementById('randomVariableForm');
            if (form.checkValidity() === false) {
                return;
            }
            Excel.run(
                context => {
                    let workbook = context.workbook;
                    let cell = getFirstSelectedCell(workbook);
                    return context.sync()
                        .then(() => {
                            const distributionModel = getSelectedDistributionModel();
                            const randomVariable = createRandomVariable(distributionModel, cell);
                            const isRandomVariableAdded = montecarlo.addVariable(randomVariable);
                            if (isRandomVariableAdded) {
                                addComments(randomVariable, workbook, cell);
                                setRangeFormat(cell, "#3F9570", "#FFFFFF");
                            }
                        })
                }).catch(error => {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            });
        });
    }

    /*
     Agrega una salida al Model de Montecarlo
      */
    function performMontecarloOutputVariableAddition() {
        const $parameterForm = $("#outputForm");
        $parameterForm.submit(event => {
            event.preventDefault();
            const form = document.getElementById('outputForm');
            if (form.checkValidity() === false) {
                return;
            }
            Excel.run(
                context => {
                    const workbook = context.workbook;
                    const cell = getFirstSelectedCell(workbook);
                    return context.sync()
                        .then(() => {
                            const output = createOutputVariable(cell);
                            const isOutputAdded = montecarlo.addOutput(output);
                            if (isOutputAdded) {
                                addComments(output, workbook, cell);
                                setRangeFormat(cell, "#008DDD", "#FFFFFF");
                            }
                        })
                }).catch(error => {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            });
        });
    }

    /*
    Limpia la UI
     */
    function clearWorkbook() {
        Excel.run(async context => {
            context
                .workbook
                .worksheets
                .getActiveWorksheet()
                .getRange()
                .clear();
            await context.sync();
        }).catch(function (error) {
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    /*
    Una vez que la api de Office.js este lista invocamos los métodos antes comentados
     */
    loadFormValidation();
    Office.onReady(() => {
        OfficeExtension.config.extendedErrorLogging = true;
        $(() => {
            hideFormsOnBack();
            showMontecarloForm();
            showFormOnClick($btnOutput, $outputForm);
            showFormOnClick($btnRandomVariable, $randomVariableForm);

            clearWorkbook();
            setDistributionSelectOptions();
            createDistributionParameterInputs();
            performMontecarloRandomVariableAddition();
            performMontecarloOutputVariableAddition();
        });


    });

})();