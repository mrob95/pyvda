
VERSION=$(python3 -c "import pyvda; print(pyvda.__version__)")

git tag -a v$VERSION