/***
 * La finalidad de este archivo es de definir las distribuciones a utilizar.
 */
const jStat = require('jstat').jStat;

export const distributions = {
    NORMAL: {
        name: 'Normal',
        parameters: [
            {
                name: 'mean',
                displayName: 'Mean',
                defaultValue: 0,
            },
            {
                name: 'standardDev',
                displayName: 'Std Deviation',
                defaultValue: 1,
            }

        ],
        pdf: (mean, std) => {
            return jStat.normal.sample(mean, std)
        }
    },
    UNIFORM: {
        name: 'Uniform',
        parameters: [
            {
                name: 'min',
                displayName: 'Min',
                defaultValue: 0,
            },
            {
                name: 'max',
                displayName: 'Max',
                defaultValue: 0,
            }
        ],
        pdf: (min, max) => {
            return jStat.uniform.sample(min, max)
        }
    },
    BETA: {
        name: 'Beta',
        parameters: [
            {
                name: 'alfa',
                displayName: 'Alfa',
                defaultValue: 0,
            },
            {
                name: 'beta',
                displayName: 'Beta',
                defaultValue: 0,
            }
        ],
        pdf: (alfa, beta) => {
            return jStat.beta.sample(alfa, beta)
        }

    },
    PERT: {
        name: 'Beta-Pert',
        parameters: [
            {
                name: 'optimistic',
                displayName: 'Optimistic',
                defaultValue: 0,
            },
            {
                name: 'mode',
                displayName: 'Mode',
                defaultValue: 0,
            },
            {
                name: 'pessimistic',
                displayName: 'Pessimistic',
                defaultValue: 0,
            }
        ],
        pdf: (pessimistic, mode, optimistic) => {
            const divider = optimistic - pessimistic;
            const alfa = (4 * mode + optimistic - 5 * pessimistic) / divider;
            const beta = (5 * optimistic - pessimistic - 4 * mode) / divider;
            return jStat.beta.sample(alfa, beta)
        }

    },
    EXPONENTIAL: {
        name: 'Exponential',
        parameters: [
            {
                name: 'lambda',
                displayName: 'Lambda',
                defaultValue: 0,
            }
        ],
        pdf: (lambda) => {
            return jStat.exponential.sample(lambda)
        }

    },
    DISCRETE: {
        name: 'Discrete',
        parameters: [
            {
                name: 'probability',
                displayName: 'Probability',
                defaultValue: 0,
            },
            {
                name: 'value',
                displayName: 'Value',
                defaultValue: 0,
            }
        ],
        pdf: (parameters) => {
            const rnd = Math.random();
            let prev = 0;
            let probability = []
            Array.from(parameters.entries())
                .forEach(prob => {
                    prev = prev + parseFloat(prob[1]);
                    probability.push([prob[0], prev]);
                });

            for (let i = 0; i < probability.length; i++) {
                if (rnd <= probability[i][1]) return probability[i][0];
            }

        }

    }
};