To update all Python modules, you can use pip, which is Python's package manager. Here's the command:

```
pip list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1 | xargs -n1 pip install -U
```

This command:
1. Lists all outdated packages in a freeze format
2. Filters out any editable packages (those installed with -e)
3. Extracts just the package names
4. Updates each package one by one

If you're using Python 3, you might need to use `pip3` instead of `pip`:

```
pip3 list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1 | xargs -n1 pip3 install -U
```

For a simpler alternative that works on Windows too:

```
python -m pip install --upgrade pip
pip install --upgrade $(pip list --outdated --format=freeze | cut -d = -f 1)
```

Remember to consider using a virtual environment when updating packages to avoid potential system-wide conflicts.