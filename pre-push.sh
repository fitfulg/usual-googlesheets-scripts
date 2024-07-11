#!/bin/bash

# cd to the project directory
cd "$(git rev-parse --show-toplevel)"

# execute script
./concatenate.sh

git add combined-scripts.js
