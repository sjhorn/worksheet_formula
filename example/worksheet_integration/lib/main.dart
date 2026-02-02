// ignore_for_file: avoid_print

import 'dart:async';

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
  late final StreamSubscription<DataChangeEvent> _dataSubscription;
  late final WorksheetController _controller;
  late final EditController _editController;
  late final LayoutSolver _layoutSolver;

  // For positioning the editor overlay
  Rect? _editingCellBounds;

  String _infoText = 'Tap a cell to see its details';

  final int _rowCount = 20;
  final int _columnCount = 6;

  final double _defaultRowHeight = 20.0;
  final double _defaultColumnWidth = 94.0;

  final double _headerWidth = 40.0;
  final double _headerHeight = 20.0;

  @override
  void initState() {
    super.initState();

    _engine = FormulaEngine();
    _engine.registerFunction(_DiscountFunction());

    _rawData = SparseWorksheetData(
      rowCount: _rowCount,
      columnCount: _columnCount,
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
    _controller.selectionController.addListener(_onSelectionChange);

    _editController = EditController();

    // Layout solver for cell bounds calculation
    _layoutSolver = LayoutSolver(
      rows: SpanList(count: _rowCount, defaultSize: _defaultRowHeight),
      columns: SpanList(count: _columnCount, defaultSize: _defaultColumnWidth),
    );
  }

  void _onSelectionChange() {
    final coord = _controller.selectionController.anchor!;
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
  void dispose() {
    _controller.selectionController.removeListener(_onSelectionChange);
    _dataSubscription.cancel();
    _data.dispose();
    _rawData.dispose();
    _controller.dispose();
    _editController.dispose();
    super.dispose();
  }

  void _onCellTap(CellCoordinate coord) {
    if (_editController.isEditing && _editController.editingCell != coord) {
      _editController.commitEdit(onCommit: _onCommit);
    }
  }

  void _onCellEdit(CellCoordinate cell) {
    // Calculate cell bounds for the editor overlay

    final cellLeft =
        _layoutSolver.getColumnLeft(cell.column) * _controller.zoom;
    final cellTop = _layoutSolver.getRowTop(cell.row) * _controller.zoom;
    final cellWidth =
        _layoutSolver.getColumnWidth(cell.column) * _controller.zoom;
    final cellHeight = _layoutSolver.getRowHeight(cell.row) * _controller.zoom;

    // Adjust for scroll offset and headers
    final adjustedLeft = cellLeft - _controller.scrollX + _headerWidth;
    final adjustedTop = cellTop - _controller.scrollY + _headerHeight;

    setState(() {
      _editingCellBounds = Rect.fromLTWH(
        adjustedLeft,
        adjustedTop,
        cellWidth,
        cellHeight,
      );
    });

    // Start editing
    final currentValue = _rawData.getCell(cell);
    _editController.startEdit(
      cell: cell,
      currentValue: currentValue,
      trigger: EditTrigger.doubleTap,
    );
  }

  void _onCommit(CellCoordinate cell, CellValue? value) {
    setState(() {
      _data.setCell(cell, value);
      _editingCellBounds = null;
    });
  }

  void _onCancel() {
    setState(() {
      _editingCellBounds = null;
    });
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
            child: Text(
              _infoText,
              style: Theme.of(context).textTheme.bodyMedium,
            ),
          ),
          Expanded(
            child: Stack(
              children: [
                WorksheetTheme(
                  data: WorksheetThemeData(
                    rowHeaderWidth: _headerWidth,
                    columnHeaderHeight: _headerHeight,
                    defaultRowHeight: _defaultRowHeight,
                    defaultColumnWidth: _defaultColumnWidth,
                  ),
                  child: Worksheet(
                    data: _data,
                    rowCount: _rowCount,
                    columnCount: _columnCount,
                    controller: _controller,
                    onCellTap: _onCellTap,
                    onEditCell: _onCellEdit,
                  ),
                ),

                // Cell editor overlay
                if (_editController.isEditing && _editingCellBounds != null)
                  CellEditorOverlay(
                    editController: _editController,
                    cellBounds: _editingCellBounds!,
                    onCommit: _onCommit,
                    onCancel: _onCancel,
                  ),
              ],
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
