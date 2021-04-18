pushall:
	git push && find . -maxdepth 1 -type d \( ! -name . ! -name .git \) -exec bash -c "cd '{}' && clasp push" \;

push:
	git push && cd ${dir} && clasp push && cd ..

pull:
	git pull && cd ${dir} && clasp pull && cd ..

pushc:
	cd ${dir} && clasp push && cd ..

pullc:
	cd ${dir} && clasp pull && cd ..

code:
	echo Opening script in ${dir}
	cd ${dir} && code *.js || code *.gs && cd ..

clone:
	mkdir ${dir} && cd ${dir} && clasp clone $(id) && cd ..
