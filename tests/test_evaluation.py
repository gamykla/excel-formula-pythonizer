from nose.tools import eq_
from formulas import evaluation


def test_is_formula():
    yield eq_, evaluation.is_formula("=1 + 2"), True
    yield eq_, evaluation.is_formula("= 1 + 2"), True
    yield eq_, evaluation.is_formula(" = 1 + 2 "), True
    yield eq_, evaluation.is_formula("= A1 + 2 "), True
    yield eq_, evaluation.is_formula("=1"), True
    yield eq_, evaluation.is_formula("= A1 / A2 + A3 - A4"), True
    yield eq_, evaluation.is_formula("=sum(E84:E86)"), True
    yield eq_, evaluation.is_formula("2"), False
    yield eq_, evaluation.is_formula("Hello"), False


def test_evaluate_basic_formula():
    def excel_value_provider(excel_cell_name):
        return 1

    evaluator = evaluation.FormulaEvaluator(excel_value_provider)

    yield eq_, evaluator.evaluate_formula("=1"), 1
    yield eq_, evaluator.evaluate_formula("=1 / 2.0"), 0.5
    yield eq_, evaluator.evaluate_formula("=1+1"), 2
    yield eq_, evaluator.evaluate_formula("=1+A1"), 2
    yield eq_, evaluator.evaluate_formula("=1 + 'Dummy Data'!B2 + A1"), 3
    yield eq_, evaluator.evaluate_formula("='Dummy Data'!A1 * 'Dummy Data'!B2 + A1 + A2 + A3"), 4


def test_recursive_evaluate_formula_that_is_referencing_formula():

    def excel_value_provider(excel_cell_name):
        if excel_cell_name == "A1":
            return "=A3 * 1"
        if excel_cell_name == "A3":
            return 2

    evaluator = evaluation.FormulaEvaluator(excel_value_provider)
    yield eq_, evaluator.evaluate_formula("=1 + A1"), 3


def test_evaluate_builtin_functions():

    def excel_value_provider(name):
        if name == "F37":
            return 1

    evaluator = evaluation.FormulaEvaluator(excel_value_provider)
    #yield eq_, evaluator.evaluate_formula("=F37+sum(F38:F39)"), 2
