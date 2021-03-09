#!/usr/bin/env bash
set -x

# to add input checking.
RAW_TAG=$1

# get folder path to zip.

if [[ ${RAW_TAG} == *"botJsDefault"* ]];then
    cd ./bot/js
    zip -r ../../bot_JavaScript_default_${RAW_TAG}.zip .
    cd ../..
elif [[ ${RAW_TAG} == *"botTsDefault"* ]];then
    cd ./bot/ts
    zip -r ../../bot_TypeScript_default_${RAW_TAG}.zip .
    cd ../..
else # botCsDefault
    cd ./bot/csharp
    zip -r ../../bot_CSharp_default_${RAW_TAG}.zip .
    cd ../..
fi
