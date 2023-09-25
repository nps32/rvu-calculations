# List all files in the current directory with the .xlsx file ending
# Initialize an empty string to store the output
output=""

# List all files in the current directory with the .xlsx file ending
for file in *.xlsx; do
    # Append the filename to the output string
    output="${output}${file}\n"
done

# Return the output
echo -e "$output"
