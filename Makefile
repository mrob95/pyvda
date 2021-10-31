test:
	python3 -m pytest --cov-report term-missing --cov=pyvda tests/

clean:
	rm -rf dist build

dist: clean
	python setup.py sdist bdist_wheel

package: dist
	twine upload dist/*

tag:
	bash create_tag.sh
