.PHONY: setup extract ask eval report all

setup:
	python -m pip install -U pip
	pip install -e ..
	pip install -e .

extract:
	exbench extract --case all --method all

ask:
	exbench ask --case all --method all --model gpt-4o

eval:
	exbench eval --case all --method all

report:
	exbench report

all: extract ask eval report
