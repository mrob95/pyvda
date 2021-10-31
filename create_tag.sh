
VERSION=$(python -c "import pyvda; print(pyvda.__version__)")

git tag v$VERSION
git push origin v$VERSION
