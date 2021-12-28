# transparent_correction

Simple code for a more transparent, criteria-based correction of exams

## The basic idea:

- The correction relies on a [spreadsheet correction template](https://github.com/majeau-bettez/transparent_correction/tree/main/template) and a Python grade calculator
- The corrector defines "correction codes" that correspond to the criteria that are tested in the exam, or to mistakes that show up during the correction
- The corrector then decides the relative weight of these criteria  or mistakes
- The template is run and we get the grades

## The advantages:

- Save time by defining clearly each error _once_ in the template, and then write only short error codes on each exam copy.
- Dissociate the identification of mistakes (objective) and the evaluation of their importance (subjective). It is easier to do the latter in an iterative manner _after_ having gained the overview of entire exam pool.
- Get an overview of the performance of the class (histograms, etc.)
- Allow for an efficient, detailed, private, transparent, and unambiguous feedback to the students, with personnalized email reports.


## Main issues:

- The template is only available in French for now.
- Limitted documentation

## Best explained by a simple demo

[See this demo](https://github.com/majeau-bettez/transparent_correction/blob/main/demo/demo_correction_basic.ipynb) with [the associated demo template](https://github.com/majeau-bettez/transparent_correction/blob/main/demo/demo_correction_template.xlsx) for an overview of the workflow and the basic features.



