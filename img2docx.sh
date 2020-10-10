
shell_dir="$0"
if [ -L "$0" ]
then 
	shell_dir=$(dirname $(readlink -f "$0"))
else
	shell_dir=$(dirname "$0")
fi

if !(printf '\033[8m' && (pip3 list | grep docx)) then
	pip3 install --user python-docx
fi
printf '\033[m'
/bin/python3 "$shell_dir/img2docx.py" "$@"
