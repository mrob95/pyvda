test:
	python3 -m pytest --cov-report term-missing --cov=pyvda tests/

clean:
	rm -r dist build

dist: clean
	python3 setup.py sdist bdist_wheel

package: dist
	twine upload dist/*

tag:
	bash create_tag.sh