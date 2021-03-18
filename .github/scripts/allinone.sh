#!/usr/bin/env bash
set -x

# to add input checking.
RAW_TAG=$1

# get folder path to zip.
if [[ ${RAW_TAG} == *"tab.JavaScript.default"* ]];then
    cd ./tab/js/default
    zip -r ../../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"tab.JavaScript.with-function"* ]];then
    cd ./tab/js/with-function
    zip -r ../../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"function-base.JavaScript.default"* ]];then
    cd ./function-base/js/default
    zip -r ../../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"function-trigger.JavaScript.HTTPTrigger"* ]];then
    cd ./function-trigger/js/HTTPTrigger
    zip -r ../../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"bot.JavaScript.default"* ]];then
    cd ./bot/js
    zip -r ../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"bot.TypeScript.default"* ]];then
    cd ./bot/ts
    zip -r ../../${RAW_TAG}.zip .
else # botCsDefault
    cd ./bot/csharp
    zip -r ../../${RAW_TAG}.zip .
fi

cd ../..
