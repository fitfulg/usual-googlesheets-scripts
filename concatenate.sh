#!/bin/bash

output_file="concat-script.js"
echo "// Auto-generated file with all JS scripts" > $output_file

find . -name "*.js" -not -name "$output_file" -not -name "jest.config.cjs" -not -name "eslint.config.cjs" -not -path "./tests/*" -not -path "./node_modules/*" | while read file; do
  echo -e "\n// Contents of $file\n" >> $output_file
  sed '/if (typeof module !== '"'"'undefined'"'"' && module\.exports)/,/^}$/d; /\/\* eslint-disable no-unused-vars \*\//d' "$file" >> $output_file
done

echo "All JavaScript files have been concatenated into $output_file"
