# http://stackoverflow.com/questions/5049489/evaluating-mathematical-expressions-in-python
# http://eli.thegreenplace.net/2009/11/28/python-internals-working-with-python-asts

import ast

e0 = "1010110"
e1 = "=A1 * A2"
e2 = "=A1 / A2"
e3 = "=A1 + A2"

e5 = "=sum(E84:E86)"
e6 = "=sumif('Dummy Data'!$AO$5:$AO$42,B86,'Dummy Data'!$AN$5:$AN$42)"


class MyTransformer(ast.NodeTransformer):
    def visit_Name(self, node):
        return ast.Num(1)


source = "1 + A4"
node = ast.parse(source, mode='eval')
print ast.dump(node)
MyTransformer().visit(node)
print ast.dump(node)
node = ast.fix_missing_locations(node)
print eval(compile(node, '<string>', mode='eval'))

