#!/bin/bash

resolve_7z() {
    if [ -n "$SEVEN_ZIP" ] && [ -f "$SEVEN_ZIP" ]; then
        echo "$SEVEN_ZIP"
        return 0
    fi

    local candidates=(
        "/c/Program Files/7-Zip/7z.exe"
        "/mnt/c/Program Files/7-Zip/7z.exe"
    )

    local path
    for path in "${candidates[@]}"; do
        if [ -f "$path" ]; then
            echo "$path"
            return 0
        fi
    done

    if command -v 7z >/dev/null 2>&1; then
        command -v 7z
        return 0
    fi

    return 1
}

read_password() {
    if [ ! -f .secret ]; then
        return 1
    fi
    tr -d '\r\n' < .secret
}
