"""Read simple uncached Excel volume formulas without running workbook code."""
import ast
import re
from decimal import Decimal, ROUND_HALF_UP
from openpyxl.utils.cell import range_boundaries, get_column_letter

class FormulaValues:
    def __init__(self, workbook, cached):
        self.workbook=workbook;self.cached=cached;self.memo={};self.visiting=set()

    def value(self,sheet,address):
        address=address.replace('$','').upper();key=(sheet,address)
        if key in self.memo: return self.memo[key]
        if key in self.visiting or len(self.visiting)>80: raise ValueError('Циклическая формула')
        cached=self.cached[sheet][address].value
        if cached is not None:
            result=self.numeric(cached);self.memo[key]=result;return result
        original=self.workbook[sheet][address].value
        if original is None: return Decimal(0)
        self.visiting.add(key)
        try:
            result=self.evaluate(str(original)[1:],sheet) if isinstance(original,str) and original.startswith('=') else self.numeric(original)
            if not result.is_finite() or abs(result)>Decimal('1e18'): raise ValueError('Недопустимый результат формулы')
            self.memo[key]=result;return result
        finally: self.visiting.remove(key)

    def numeric(self,value):
        if isinstance(value,bool): raise ValueError('Логическое значение вместо объёма')
        try: return Decimal(str(value).replace(' ','').replace('\u00a0','').replace(',','.'))
        except Exception: raise ValueError('Формула с нечисловым значением')

    def evaluate(self,expression,sheet):
        if '[' in expression or '!' in expression or '"' in expression or len(expression)>2000:
            raise ValueError('Формула требует пересчёта в Excel')
        # Excel ranges and references become data arguments in a restricted AST.
        expression=re.sub(r'(?i)\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)',lambda m:'RANGE("'+''.join(m.group(1,2))+':'+''.join(m.group(3,4))+'")',expression)
        expression=re.sub(r'(?i)(?<![\w"\:])\$?([A-Z]{1,3})\$?(\d+)(?![\w"\:])',lambda m:'CELL("'+''.join(m.group(1,2))+'")',expression)
        expression=re.sub(r'(\d+(?:\.\d+)?)%',r'(\1/100)',expression).replace(';',',')
        tree=ast.parse(expression,mode='eval')
        if sum(1 for _ in ast.walk(tree))>400: raise ValueError('Слишком сложная формула')
        def flatten(items):
            return [v for item in items for v in (item if isinstance(item,list) else [item])]
        def calc(node):
            if isinstance(node,ast.Constant):
                if isinstance(node.value,(int,float)): return Decimal(str(node.value))
                if isinstance(node.value,str): return node.value
            if isinstance(node,ast.UnaryOp) and isinstance(node.op,(ast.UAdd,ast.USub)):
                value=calc(node.operand);return value if isinstance(node.op,ast.UAdd) else -value
            if isinstance(node,ast.BinOp):
                left,right=calc(node.left),calc(node.right)
                if not isinstance(left,Decimal) or not isinstance(right,Decimal): raise ValueError('Нечисловое выражение')
                if isinstance(node.op,ast.Add): return left+right
                if isinstance(node.op,ast.Sub): return left-right
                if isinstance(node.op,ast.Mult): return left*right
                if isinstance(node.op,ast.Div): return left/right
            if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and not node.keywords:
                name=node.func.id.upper();args=[calc(a) for a in node.args]
                if name=='CELL' and len(args)==1: return self.value(sheet,args[0])
                if name=='RANGE' and len(args)==1:
                    x1,y1,x2,y2=range_boundaries(args[0])
                    if (x2-x1+1)*(y2-y1+1)>10000: raise ValueError('Слишком большой диапазон')
                    return [self.value(sheet,f'{get_column_letter(x)}{y}') for y in range(y1,y2+1) for x in range(x1,x2+1)]
                values=flatten(args)
                if not all(isinstance(v,Decimal) for v in values): raise ValueError('Нечисловой аргумент')
                if name=='SUM': return sum(values,Decimal(0))
                if name=='AVERAGE' and values: return sum(values,Decimal(0))/len(values)
                if name=='MIN' and values: return min(values)
                if name=='MAX' and values: return max(values)
                if name=='ABS' and len(values)==1: return abs(values[0])
                if name=='ROUND' and len(values)==2 and values[1]==values[1].to_integral_value() and abs(values[1])<=12:
                    return values[0].quantize(Decimal(10)**-int(values[1]),rounding=ROUND_HALF_UP)
            raise ValueError('Формула требует пересчёта в Excel')
        result=calc(tree.body)
        if not isinstance(result,Decimal): raise ValueError('Формула без числового результата')
        return result
