#!/usr/bin/env bash
set -x
git add .
git commit -m "release: $1"
git push origin main
git tag $1
git push origin $1

