#!/bin/bash

PEOPLE_FILE="templates/.people.xlsx"

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

has_archive_sources() {
    if [ -d "Data" ] && [ -n "$(ls -A Data 2>/dev/null)" ]; then
        return 0
    fi
    if [ -f "$PEOPLE_FILE" ]; then
        return 0
    fi
    return 1
}

needs_unpack() {
    if [ ! -d "Data" ] || [ -z "$(ls -A Data 2>/dev/null)" ]; then
        return 0
    fi
    if [ ! -f "$PEOPLE_FILE" ]; then
        return 0
    fi
    return 1
}

pack_encrypted_archive() {
    local seven_zip=$1
    local password=$2
    local -a paths=()

    if [ -d "Data" ] && [ -n "$(ls -A Data 2>/dev/null)" ]; then
        paths+=("./Data/")
    fi
    if [ -f "$PEOPLE_FILE" ]; then
        paths+=("./$PEOPLE_FILE")
    fi

    rm -f data.7z
    "$seven_zip" a -t7z -mhe=on -p"$password" data.7z "${paths[@]}"
}

fix_people_file_location() {
    if [ -f ".people.xlsx" ] && [ ! -f "$PEOPLE_FILE" ]; then
        mkdir -p templates
        mv -f ".people.xlsx" "$PEOPLE_FILE"
    elif [ -f ".people.xlsx" ]; then
        rm -f ".people.xlsx"
    fi
}

unpack_encrypted_archive() {
    local seven_zip=$1
    local password=$2

    mkdir -p Data templates
    "$seven_zip" x -p"$password" -o. data.7z -y
    fix_people_file_location
}

unstage_local_secrets() {
    git restore --staged Data/ "$PEOPLE_FILE" > /dev/null 2>&1 \
        || git reset HEAD -- Data/ "$PEOPLE_FILE" > /dev/null 2>&1
}
