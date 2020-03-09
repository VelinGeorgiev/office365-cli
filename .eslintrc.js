module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: [
    '@typescript-eslint',
  ],
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/eslint-recommended',
    'plugin:@typescript-eslint/recommended',
  ],
  overrides: [
    {
      files: ["*.ts"],
      rules: {
        "@typescript-eslint/no-explicit-any": "off",
        "@typescript-eslint/no-var-requires": "off",
        "@typescript-eslint/explicit-function-return-type": "off",
        "@typescript-eslint/no-empty-function": "off",
        "@typescript-eslint/no-inferrable-types": "off",
        "@typescript-eslint/camelcase": "off",
        "@typescript-eslint/no-unused-vars": "off",
        "@typescript-eslint/type-annotation-spacing":"off",
        "@typescript-eslint/adjacent-overload-signatures":"off",
        "@typescript-eslint/member-delimiter-style":"off",
        "@typescript-eslint/consistent-type-assertions":"off",
        "no-useless-escape":"off",
        "@typescript-eslint/no-use-before-define":"off",
        "@typescript-eslint/ban-types":"off",
        "@typescript-eslint/no-empty-interface":"off",
        "@typescript-eslint/interface-name-prefix":"off",
        "@typescript-eslint/prefer-namespace-keyword":"off",
        "@typescript-eslint/no-namespace":"off",
        "no-empty":"off",

        "eqeqeq":"warn",
      }
    },
    {
      files: ["*.spec.ts"],
      rules: {
        "@typescript-eslint/no-empty-function": "off",
        "no-useless-escape":"off",
        "prefer-const":"off",
        "no-useless-catch":"off"
      }
    },
    {
      files: ["**/project-upgrade/**/*.ts","**/project-upgrade.ts"],
      rules: {
        "@typescript-eslint/class-name-casing":"off",
        "no-extra-semi":"off",
      }
    }
  ]
};