module.exports = [
    {
        ignores: ['node_modules/**', 'dist/**'],
    },
    {
        files: ['**/*.js'],
        rules: {
            'no-unused-vars': 'error',
            'no-console': 'warn',
        },
    },
];
