#!/bin/bash

output_file="concat-script.js"
echo "// Auto-generated file with all JS scripts" > $output_file

find . -name "*.js" -not -name "$output_file" | while read file; do
  echo -e "\n// Contents of $file\n" >> $output_file
  cat "$file" >> $output_file
done

echo "All JavaScript files have been concatenated into $output_file"
