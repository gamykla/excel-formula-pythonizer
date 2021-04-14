import ast
import logging

import re

logger = logging.getLogger(__name__)

CELL_REFERENCE_VAR_PREFIX = "__cell_ref__"
EXCEL_FUNCTION_REFERENCE_VAR_PREFIX = "__excel_func__"


def is_formula(value):
    return str(value).strip().startswith('=')


def remove_equals(formula_string):
    return formula_string.strip().replace('=', '')


class FormulaPythonizer(object):

    def __init__(self):
        self.python_variable_name_to_excel_variable_name = {}

    def get_excel_reference_name(self, variable_name):
        return self.python_variable_name_to_excel_variable_name[variable_name]

    def convert_excel_cell_references_to_python_variables(self, formula):
        references = re.findall("\'[A-Za-z ]+\'\![A-Za-z]+[0-9]+", formula)

        for excel_variable_name in references:
            python_name = excel_variable_name
            python_name = python_name.replace("'", "")
            python_name = python_name.replace(" ", "_")
            python_name = python_name.replace("!", '_')
            python_name = CELL_REFERENCE_VAR_PREFIX + python_name

            self.python_variable_name_to_excel_variable_name[python_name] = excel_variable_name
            formula = formula.replace(excel_variable_name, python_name)

        return formula

    def get_excel_function_value(self, variable):
        pass


class FormulaVariableNameToValueTransformer(ast.NodeTransformer):

    def __init__(self, formula_pythonizer, excel_value_provider):
        super(FormulaVariableNameToValueTransformer, self).__init__()
        self.formula_pythonizer = formula_pythonizer
        self.excel_value_provider = excel_value_provider

    def visit_Name(self, node):
        return self._replace_ast_variable_name_node_with_concrete_value(node)

    def _replace_ast_variable_name_node_with_concrete_value(self, node):
        formula_variable_name = str(node.id)

        value = self._formula_variable_value(formula_variable_name)
        return self._get_ast_representation(value)

    def _formula_variable_value(self, variable):
        excel_cell_lookup_name = variable

        # the variable in the formula is a standin for an excel function
        if self._is_python_formula_variable_representing_excel_function(variable):
            return self.formula_pythonizer.get_excel_function_value(variable)

        # the variable in the formula is a standin for a direct cell lookup
        if self._is_python_formula_variable_substituting_excel_cell_reference(variable):
            excel_cell_lookup_name = self.formula_pythonizer.get_excel_reference_name(variable)

        return self._get_excel_value(excel_cell_lookup_name)

    def _get_ast_representation(self, value):
        if isinstance(value, (int, long, float, complex)):
            return ast.Num(value)
        else:
            raise Exception("Can't find an ast type associated with {}".format(value))

    def _is_python_formula_variable_substituting_excel_cell_reference(self, variable):
        # check if the variable is in place of an excel reference like 'Dummy Data'!AN5:AN42 that doesn't work in AST
        return variable.startswith(CELL_REFERENCE_VAR_PREFIX)

    def _is_python_formula_variable_representing_excel_function(self, variable):
        return variable.startswith(EXCEL_FUNCTION_REFERENCE_VAR_PREFIX)

    def _get_excel_value(self, excel_lookup_name):
        excel_value = self.excel_value_provider(excel_lookup_name)

        if is_formula(excel_value):
            formula_evaluator = FormulaEvaluator(self.excel_value_provider)
            excel_value = formula_evaluator.evaluate_formula(excel_value)

        return excel_value


class FormulaEvaluator(object):

    def __init__(self, excel_value_provider):
        self.excel_value_provider = excel_value_provider

    def evaluate_formula(self, formula):
        if not is_formula(formula):
            raise Exception(str(formula) + " is not a formula")

        logger.debug("Source formula is {}".format(formula))

        formula = remove_equals(formula)

        pythonizer = FormulaPythonizer()
        formula = pythonizer.convert_excel_cell_references_to_python_variables(formula)

        node = ast.parse(formula, mode='eval')

        logger.debug("formula ast is {}".format(ast.dump(node)))

        FormulaVariableNameToValueTransformer(pythonizer, self.excel_value_provider).visit(node)

        logger.debug("forumala with variable substitution is {}".format(ast.dump(node)))

        node = ast.fix_missing_locations(node)
        codeobj = compile(node, '<string>', mode='eval')
        return eval(codeobj)
