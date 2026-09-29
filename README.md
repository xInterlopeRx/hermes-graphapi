# Hermes Microsoft Graph API extension

This repository contains the app-only Microsoft Graph endpoint package and
compatibility imports for Hermes. The installer copies only
`gateway/msgraph_enterprise/` and
`gateway/platforms/msgraph_enterprise/` into an existing Hermes agent checkout.
It does not modify `.env` or other Hermes files.

## Required Microsoft Graph permissions

This extension requires the following app-only Microsoft Graph permissions:

- `Calendars.ReadWrite`
- `Contacts.ReadWrite`
- `Mail.ReadWrite`
- `Mail.Send`

`Mail.ReadWrite` is specifically required to create, update, and delete drafts
and messages. `Mail.Send` is only required for the `send_message` operation.

## Install

Run from a shell with Git, Python, pytest, and the Graph package's Python
dependencies available. The default Hermes root is
`${HERMES_ROOT:-${HERMES_HOME:-$HOME/.hermes}/hermes-agent}`; it must already
contain `gateway/` and `gateway/platforms/`.

```sh
curl -fsSL https://raw.githubusercontent.com/xInterlopeRx/hermes-graphapi/main/install.sh | bash
```

Set `GRAPHAPI_REF` to a branch or tag to install that version. `HERMES_ROOT`
overrides the root directly; otherwise `HERMES_HOME` selects the Hermes home
directory. For example:

```sh
export GRAPHAPI_REF=v1.0.0
export HERMES_HOME="$HOME/.hermes"
curl -fsSL https://raw.githubusercontent.com/xInterlopeRx/hermes-graphapi/main/install.sh | bash
```

Before replacing packages, the installer saves any existing Graph packages in
a timestamped `.graphapi-backup-*` directory inside the Hermes root. Re-running
the installer replaces only those two Graph package directories. It compiles
the installed packages and runs this repository's focused tests against the
installed copy; any install or verification failure triggers a rollback from
the saved backup. Installations fail closed when expected directories,
commands, packages, or tests are unavailable.

## Focused tests

The test-only `agent.secret_scope` shim returns defaults and requires neither
the Hermes source tree nor credentials. Run with:

```sh
python -m pip install pytest requests
python -m pytest -q
```

Tests use mocked HTTP sessions and never contact Microsoft Graph.
