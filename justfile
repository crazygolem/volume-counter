set positional-arguments
set dotenv-load

# List the recipes
@default:
    just --list

# Run clasp commands
clasp *args:
    bun run clasp "$@"

# Generate the clasp configuration using variables from the environment or an .env file
init-clasp:
    #!/usr/bin/env bun

    const config = {
        // The script ID can be found in the Apps Script IDE under project settings
        "scriptId": process.env.GAPPS_SCRIPTID
            ?? (() => { throw new Error("GAPPS_SCRIPTID env var missing") })(),
        "rootDir": "src",
        "scriptExtensions": [
            ".js",
            ".gs"
        ],
        "htmlExtensions": [
            ".html"
        ],
        "jsonExtensions": [
            ".json"
        ],
        "filePushOrder": [],
        "skipSubdirectories": false
    }

    await Bun.write('.clasp.json', JSON.stringify(config, null, 2))
