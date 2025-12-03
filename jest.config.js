module.exports = {
    testEnvironment: 'node',
    testMatch: ['**/*.test.js'],
    transform: {
        '^.+\\.js$': 'babel-jest'
    },
    transformIgnorePatterns: [
        '/node_modules/'
    ],
    moduleFileExtensions: ['js'],
    verbose: true
};
