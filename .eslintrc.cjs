/**
 * @type {import('eslint').Linter.Config}
 */
module.exports = {
  extends: '@zzavidd/eslint-config/node-ts',
  parserOptions: {
    sourceType: 'module',
    tsconfigRootDir: __dirname,
    project: ['tsconfig.json', 'projects/*/tsconfig.json'],
  },
  settings: {
    'import/resolver': {
      typescript: {
        project: ['tsconfig.json', 'projects/**/tsconfig.json'],
      },
    },
  },
};
