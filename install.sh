#!/usr/bin/env bash
set -Eeuo pipefail

GRAPHAPI_REF=${GRAPHAPI_REF:-main}
HERMES_ROOT=${HERMES_ROOT:-${HERMES_HOME:-$HOME/.hermes}/hermes-agent}
REPOSITORY_URL=https://github.com/xInterlopeRx/hermes-graphapi.git

if [[ -z "$GRAPHAPI_REF" ]]; then
    printf '%s\n' 'GRAPHAPI_REF must not be empty.' >&2
    exit 1
fi

for command_name in git python; do
    if ! command -v "$command_name" >/dev/null 2>&1; then
        printf 'Required command not found: %s\n' "$command_name" >&2
        exit 1
    fi
done

if [[ ! -d "$HERMES_ROOT" ]]; then
    printf 'Hermes root does not exist: %s\n' "$HERMES_ROOT" >&2
    exit 1
fi
HERMES_ROOT=$(cd "$HERMES_ROOT" && pwd -P)
if [[ ! -d "$HERMES_ROOT/gateway" || -L "$HERMES_ROOT/gateway" || \
      ! -d "$HERMES_ROOT/gateway/platforms" || -L "$HERMES_ROOT/gateway/platforms" ]]; then
    printf 'Hermes root is missing expected gateway directories: %s\n' "$HERMES_ROOT" >&2
    exit 1
fi

TEMP_DIR=$(mktemp -d)
STAGE_DIR=
BACKUP_DIR=
INSTALL_STARTED=0
INSTALL_SUCCEEDED=0
HAD_CANONICAL_PACKAGE=0
HAD_COMPAT_PACKAGE=0

restore_package() {
    local relative_path=$1
    local was_present=$2
    local target="$HERMES_ROOT/$relative_path"
    rm -rf -- "$target"
    if [[ "$was_present" == 1 ]]; then
        mkdir -p -- "$(dirname "$target")"
        cp -a -- "$BACKUP_DIR/$relative_path" "$target"
    fi
}

finish() {
    local status=$?
    local rollback_failed=0
    trap - EXIT
    if [[ "$INSTALL_STARTED" == 1 && "$INSTALL_SUCCEEDED" != 1 ]]; then
        printf '%s\n' 'Install or verification failed; restoring the previous Graph packages.' >&2
        if ! restore_package gateway/msgraph_enterprise "$HAD_CANONICAL_PACKAGE"; then
            rollback_failed=1
        fi
        if ! restore_package gateway/platforms/msgraph_enterprise "$HAD_COMPAT_PACKAGE"; then
            rollback_failed=1
        fi
        if [[ "$rollback_failed" == 1 ]]; then
            printf 'Automatic rollback was incomplete. The saved backup is at: %s\n' "$BACKUP_DIR" >&2
            status=1
        else
            printf 'Previous Graph packages restored from: %s\n' "$BACKUP_DIR" >&2
        fi
    fi
    [[ -z "$STAGE_DIR" ]] || rm -rf -- "$STAGE_DIR"
    rm -rf -- "$TEMP_DIR"
    exit "$status"
}
trap finish EXIT
trap 'exit 1' HUP INT TERM

CHECKOUT="$TEMP_DIR/hermes-graphapi"
printf 'Fetching %s at ref %s...\n' "$REPOSITORY_URL" "$GRAPHAPI_REF"
git clone --depth 1 --branch "$GRAPHAPI_REF" "$REPOSITORY_URL" "$CHECKOUT"

for relative_path in gateway/msgraph_enterprise gateway/platforms/msgraph_enterprise; do
    if [[ ! -d "$CHECKOUT/$relative_path" || -L "$CHECKOUT/$relative_path" ]]; then
        printf 'Clone is missing expected Graph package directory: %s\n' "$relative_path" >&2
        exit 1
    fi
done

BACKUP_DIR="$HERMES_ROOT/.graphapi-backup-$(date -u +%Y%m%dT%H%M%SZ)-$$"
mkdir -- "$BACKUP_DIR"
for relative_path in gateway/msgraph_enterprise gateway/platforms/msgraph_enterprise; do
    destination="$HERMES_ROOT/$relative_path"
    if [[ -L "$destination" ]]; then
        printf 'Refusing to replace a symlink: %s\n' "$destination" >&2
        exit 1
    fi
    if [[ -e "$destination" && ! -d "$destination" ]]; then
        printf 'Refusing to replace a non-directory Graph path: %s\n' "$destination" >&2
        exit 1
    fi
    if [[ -d "$destination" ]]; then
        mkdir -p -- "$BACKUP_DIR/$(dirname "$relative_path")"
        cp -a -- "$destination" "$BACKUP_DIR/$relative_path"
        if [[ "$relative_path" == gateway/msgraph_enterprise ]]; then
            HAD_CANONICAL_PACKAGE=1
        else
            HAD_COMPAT_PACKAGE=1
        fi
    fi
done

STAGE_DIR=$(mktemp -d "$HERMES_ROOT/.graphapi-stage.XXXXXX")
mkdir -p -- "$STAGE_DIR/gateway/platforms"
cp -a -- "$CHECKOUT/gateway/msgraph_enterprise" "$STAGE_DIR/gateway/"
cp -a -- "$CHECKOUT/gateway/platforms/msgraph_enterprise" "$STAGE_DIR/gateway/platforms/"

INSTALL_STARTED=1
rm -rf -- "$HERMES_ROOT/gateway/msgraph_enterprise" \
    "$HERMES_ROOT/gateway/platforms/msgraph_enterprise"
mv -- "$STAGE_DIR/gateway/msgraph_enterprise" "$HERMES_ROOT/gateway/"
mv -- "$STAGE_DIR/gateway/platforms/msgraph_enterprise" "$HERMES_ROOT/gateway/platforms/"

printf '%s\n' 'Compiling installed Graph packages...'
python - "$HERMES_ROOT/gateway/msgraph_enterprise" \
    "$HERMES_ROOT/gateway/platforms/msgraph_enterprise" <<'PY'
import pathlib
import sys

for package in map(pathlib.Path, sys.argv[1:]):
    for source in package.rglob("*.py"):
        compile(source.read_bytes(), str(source), "exec")
PY

if ! python -m pytest --version >/dev/null 2>&1; then
    printf '%s\n' 'pytest is required to verify the install; install pytest and retry.' >&2
    exit 1
fi

printf '%s\n' 'Running focused Graph API tests against the installed packages...'
(
    cd "$HERMES_ROOT"
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONPATH="$CHECKOUT/tests/support:$HERMES_ROOT${PYTHONPATH:+:$PYTHONPATH}" \
        python -m pytest -q -p no:cacheprovider --import-mode=importlib \
        "$CHECKOUT/tests/gateway"
)

INSTALL_SUCCEEDED=1
printf 'Graph API installed and verified. Existing Graph files are backed up at: %s\n' "$BACKUP_DIR"
