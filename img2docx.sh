
shell_dir="$0"
if [ -L "$0" ]
then 
	shell_dir=$(dirname $(readlink -f "$0"))
else
	shell_dir=$(dirname "$0")
fi

/a/bin/python3 "$shell_dir/img2docx.py" "$@"
