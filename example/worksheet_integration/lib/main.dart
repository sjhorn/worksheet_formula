// ignore_for_file: avoid_print

import 'package:flutter/material.dart';
import 'package:worksheet/worksheet.dart';
import 'package:worksheet_formula/worksheet_formula.dart';

import 'src/formula_worksheet_data.dart';

void main() {
  runApp(const WorksheetFormulaApp());
}

class WorksheetFormulaApp extends StatelessWidget {
  const WorksheetFormulaApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Worksheet Formula Demo',
      theme: ThemeData(colorSchemeSeed: Colors.blue, useMaterial3: true),
      home: const WorksheetFormulaDemo(),
    );
  }
}

class WorksheetFormulaDemo extends StatefulWidget {
  const WorksheetFormulaDemo({super.key});

  @override
  State<WorksheetFormulaDemo> createState() => _WorksheetFormulaDemoState();
}

class _WorksheetFormulaDemoState extends State<WorksheetFormulaDemo> {
  late final FormulaEngine _engine;
  late final SparseWorksheetData _rawData;
  late final FormulaWorksheetData _data;
  late final WorksheetController _controller;
  String _infoText = 'Tap a cell to see its details';

  @override
  void initState() {
    super.initState();

    _engine = FormulaEngine();
    _engine.registerFunction(_DiscountFunction());

    _rawData = SparseWorksheetData(
      rowCount: 20,
      columnCount: 6,
      cells: {
        // Headers (row 0)
        (0, 0): 'Item'.cell,
        (0, 1): 'Price'.cell,
        (0, 2): 'Qty'.cell,
        (0, 3): 'Total'.cell,

        // Data rows
        (1, 0): 'Apples'.cell,
        (1, 1): Cell.number(1.50),
        (1, 2): Cell.number(10),
        (1, 3): '=B2*C2'.formula,

        (2, 0): 'Oranges'.cell,
        (2, 1): Cell.number(2.00),
        (2, 2): Cell.number(5),
        (2, 3): '=B3*C3'.formula,

        (3, 0): 'Bananas'.cell,
        (3, 1): Cell.number(0.75),
        (3, 2): Cell.number(12),
        (3, 3): '=B4*C4'.formula,

        // Summary
        (5, 2): 'Subtotal:'.cell,
        (5, 3): '=SUM(D2:D4)'.formula,

        (6, 2): 'Tax (10%):'.cell,
        (6, 3): '=D6*0.1'.formula,

        (7, 2): 'Total:'.cell,
        (7, 3): '=D6+D7'.formula,

        // Custom function
        (9, 2): 'With 15% discount:'.cell,
        (9, 3): '=DISCOUNT(D8, 0.15)'.formula,
      },
    );

    _data = FormulaWorksheetData(_rawData, engine: _engine);
    _controller = WorksheetController();
  }

  @override
  void dispose() {
    _data.dispose();
    _rawData.dispose();
    _controller.dispose();
    super.dispose();
  }

  void _onCellTap(CellCoordinate coord) {
    final rawValue = _rawData.getCell(coord);
    final displayValue = _data.getCell(coord);

    final notation = coord.toNotation();
    if (rawValue == null) {
      setState(() => _infoText = '$notation: (empty)');
      return;
    }

    if (rawValue.isFormula) {
      final formula = rawValue.rawValue as String;
      final result = displayValue?.displayValue ?? '(error)';
      setState(() => _infoText = '$notation: $formula = $result');
    } else {
      setState(() => _infoText = '$notation: ${rawValue.displayValue}');
    }
    print(_infoText);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Worksheet Formula Demo')),
      body: Column(
        children: [
          Container(
            width: double.infinity,
            padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
            color: Theme.of(context).colorScheme.surfaceContainerHighest,
            child: Text(_infoText, style: Theme.of(context).textTheme.bodyMedium),
          ),
          Expanded(
            child: WorksheetTheme(
              data: const WorksheetThemeData(),
              child: Worksheet(
                data: _data,
                rowCount: 20,
                columnCount: 6,
                controller: _controller,
                onCellTap: _onCellTap,
              ),
            ),
          ),
        ],
      ),
    );
  }
}

/// Custom function: DISCOUNT(price, rate) → price * (1 - rate)
class _DiscountFunction extends FormulaFunction {
  @override
  String get name => 'DISCOUNT';

  @override
  int get minArgs => 2;

  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    for (final v in values) {
      if (v.isError) return v;
    }
    final price = values[0].toNumber();
    final rate = values[1].toNumber();
    if (price == null || rate == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.number(price * (1 - rate));
  }
}
