{
  "name": "srm-dashboard-api",
  "version": "1.0.0",
  "description": "Azure Functions API backed by SharePoint Lists for the Nouryon SRM Dashboard",
  "main": "src/functions/*.js",
  "scripts": {
    "start": "func start",
    "test": "echo \"(add tests here)\" && exit 0"
  },
  "dependencies": {
    "@azure/functions": "^4.5.0",
    "@azure/identity": "^4.4.1",
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "isomorphic-fetch": "^3.0.0"
  },
  "engines": {
    "node": ">=18"
  }
}
