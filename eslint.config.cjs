module.exports = [
    {
        ignores: ['node_modules/**', 'dist/**', 'concat-script.js'],
    },
    {
        files: ['**/*.js'],
        rules: {
            'no-unused-vars': 'error',
            'no-console': 'warn',
        },
    },
];
