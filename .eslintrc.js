/**
 * @type {import('eslint').Linter.Config}
 */
module.exports = {
  extends: '@zzavidd/eslint-config/node-ts',
  parserOptions: {
    tsconfigRootDir: __dirname,
    project: ['**/tsconfig.json'],
  },
  settings: {
    'import/resolver': {
      typescript: {
        project: [
          'tsconfig.json',
        ],
      },
    },
  },
};
