#!/bin/bash
# save as git-gemini-commit and put in your $PATH
# Add tags to this script to customize the commit message generation
# gemini-commit.sh -t my-tag -t another-tag

git add -A
diff=$(git diff --cached)
# Add all tags to end of the commit message
# as [tagname:tagvalue] and if multiple of the same tag, add them
# as comma separated values
# e.g. [tagname:value1,value2,value3]

tagstring=""
if [ "$#" -gt 0 ]; then
  for tag in "$@"; do
    tagname=$(echo "$tag" | cut -d '=' -f 1)
    tagvalue=$(echo "$tag" | cut -d '=' -f 2-)
    if [ -z "$tagvalue" ]; then
      tagvalue="true"
    fi
    if [ -z "$tagstring" ]; then
      tagstring="[${tagname}:${tagvalue}"
    else
      tagstring="${tagstring},${tagvalue}"
    fi
  done
  tagstring=" ${tagstring}]"
else
  tagstring=""
fi

# Attempt to generate a commit message using Gemini gemini-cli
# If it fails, echo the error and exit
message=$(npx -y https://github.com/google-gemini/gemini-cli --prompt "Generate a commit message for the following changes:

\`\`\`diff
$diff
\`\`\`" | sed 's/^`\x60\x60//;s/`\x60\x60$//' | sed '/^diff/d' | sed '1d;$d'
)
if [ $? -ne 0 ]; then
  # Run a test prompt and gather error message from gemini-cli
  npx -y https://github.com/google-gemini/gemini-cli --prompt "Test prompt to check Gemini CLI installation"
  exit 1
fi

git commit -m "$message$tagstring"
