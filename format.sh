
python -m black --line-length 120 -C hntb_gen.py ./hntb
python -m isort --profile black hntb_gen.py ./hntb
python -m flake8 --config tests/flake8 hntb_gen.py ./hntb
