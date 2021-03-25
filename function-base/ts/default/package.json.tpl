{
  "name": "{{appName}}-api",
  "version": "1.0.0",
  "description": "",
  "scripts": {
    "build": "tsc",
    "watch": "tsc -w",
    "prestart": "npm run build",
    "start": "func start",
    "test": "echo \"No tests yet...\""
  },
  "dependencies": {
    "@azure/functions": "^1.2.2",
    "mods-server": "github:OfficeDev/mods-server-sdk-private-preview"
  },
  "devDependencies": {
    "typescript": "^3.3.3"
  }
}
