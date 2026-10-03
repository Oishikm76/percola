import 'dart:io';
import 'dart:math' as math;
import 'package:excel/excel.dart' as xl;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:fl_chart/fl_chart.dart';

void main() => runApp(const PercolaApp());

class PercolaApp extends StatelessWidget {
  const PercolaApp({super.key});

  @override
  Widget build(BuildContext context) {
    const seed = Color(0xFF1565C0); // deep blue
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Percola: Tubewell Designer',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(
          seedColor: seed,
          brightness: Brightness.light,
          secondary: const Color(0xFFFFB300),
        ),
        scaffoldBackgroundColor: const Color(0xFFF9FAFB),
        appBarTheme: const AppBarTheme(
          backgroundColor: Color(0xFF1565C0),
          foregroundColor: Colors.white,
          centerTitle: true,
        ),
        useMaterial3: true,
      ),
      home: const PercolaHome(),
    );
    }
}

// Define the TubewellPainter class outside the _PercolaHomeState class
class TubewellPainter extends CustomPainter {
  final double screenStart, screenEnd, sandTrapLen;

  TubewellPainter(this.screenStart, this.screenEnd, this.sandTrapLen)
      : assert(screenEnd >= 0, 'screenEnd must be initialized and non-negative');

  @override
  void paint(Canvas canvas, Size size) {
    final paintPipe = Paint()..color = Colors.grey.shade800..strokeWidth = 20..strokeCap = StrokeCap.round;
    final paintScreen = Paint()..color = Colors.green.shade600..strokeWidth = 20;
    final paintSand = Paint()..color = Colors.brown..strokeWidth = 20;

    final total = screenEnd + 30;
    double y(double z) => (z / total) * size.height;

    final housingTop = Offset(size.width / 2, 0);
    final housingBottom = Offset(size.width / 2, y(screenEnd + sandTrapLen + 30));
    canvas.drawLine(housingTop, housingBottom, paintPipe);

    final screenTop = Offset(size.width / 2, y(screenStart));
    final screenBottom = Offset(size.width / 2, y(screenEnd));
    canvas.drawLine(screenTop, screenBottom, paintScreen);

    final sandTop = Offset(size.width / 2, y(screenEnd));
    final sandBottom = Offset(size.width / 2, y(screenEnd + sandTrapLen));
    canvas.drawLine(sandTop, sandBottom, paintSand);

    textPainter(String txt, double z) {
      final tp = TextPainter(
        text: TextSpan(text: txt, style: const TextStyle(color: Colors.black, fontSize: 12)),
        textDirection: TextDirection.ltr,
      )..layout();
      tp.paint(canvas, Offset(size.width / 2 + 25, y(z) - 8));
    }

    textPainter('Screen: ${screenStart.toStringAsFixed(0)}–${screenEnd.toStringAsFixed(0)} ft', screenStart);
    textPainter('Sand Trap (${sandTrapLen.toStringAsFixed(0)} ft)', screenEnd + sandTrapLen / 2);
  }

  @override
  bool shouldRepaint(CustomPainter oldDelegate) => true;
}

// Removed duplicate TubewellPainter class definition.

class WellLayer {
  final double depthFrom, depthTo, fm, d60, d10, d30, cu;
  WellLayer({
    required this.depthFrom,
    required this.depthTo,
    required this.fm,
    required this.d60,
    required this.d10,
    required this.d30,
    required this.cu,
  });

  double get kIndex => d10 / cu;
}

class PercolaHome extends StatefulWidget {
  const PercolaHome({super.key});
  @override
  State<PercolaHome> createState() => _PercolaHomeState();
}

class _PercolaHomeState extends State<PercolaHome> {
  String? excelPath;
  List<WellLayer> layers = [];

  final _staticCtl = TextEditingController(text: '360');
  final _yearsCtl = TextEditingController(text: '20');
  final _declineCtl = TextEditingController(text: '2');
  final _drawdownCtl = TextEditingController(text: '15');

  double aquiferTop = 0, aquiferBottom = 0;
  double housingDepth = 0, screenStart = 0, screenEnd = 0, screenLen = 0, sandTrapLen = 10;
  late WellLayer governing;
  double multFactor = 0, d10Gravel = 0, d30Gravel = 0, rawSlotNo = 0, selectedSlot = 0, openingPct = 0, qLPH = 0;

  Future<void> pickExcel() async {
    final res = await FilePicker.platform.pickFiles(type: FileType.custom, allowedExtensions: ['xlsx', 'xls']);
    if (res == null) return;

    final path = res.files.single.path!;
    final bytes = File(path).readAsBytesSync();
    final book = xl.Excel.decodeBytes(bytes);

    final sheet = book.tables.values.first;
    final header = sheet.rows.first.map((c) => (c?.value ?? '').toString().trim()).toList();

    int iFrom = header.indexOf('Depth_from_ft');
    int iTo = header.indexOf('Depth_to_ft');
    int iFM = header.indexOf('FM');
    int iD60 = header.indexOf('D60_mm');
    int iD10 = header.indexOf('D10_mm');
    int iCu = header.indexOf('Cu');
    int iD30 = header.indexOf('D30_mm');

    if ([iFrom, iTo, iFM, iD60, iD10, iCu].any((i) => i < 0)) {
      // ignore: use_build_context_synchronously
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Missing one or more required columns.')),
      );
      return;
    }

    final parsed = <WellLayer>[];
    for (int r = 1; r < sheet.rows.length; r++) {
      final row = sheet.rows[r];
      double val(dynamic v) => (v is num)
          ? v.toDouble()
          : double.tryParse(v.toString()) ?? double.nan;

      final from = val(row[iFrom]?.value);
      final to = val(row[iTo]?.value);
      final fm = val(row[iFM]?.value);
      final d60 = val(row[iD60]?.value);
      final d10 = val(row[iD10]?.value);
      final cu = val(row[iCu]?.value);
      final d30 = iD30 >= 0 ? val(row[iD30]?.value) : math.sqrt(d10 * d60);
      if ([from, to, fm, d60, d10, cu, d30].any((v) => v.isNaN)) continue;

      parsed.add(WellLayer(
        depthFrom: from,
        depthTo: to,
        fm: fm,
        d60: d60,
        d10: d10,
        d30: d30,
        cu: cu,
      ));
    }

    setState(() {
      excelPath = path;
      layers = parsed;
    });

    // ignore: use_build_context_synchronously
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('Loaded ${layers.length} layers')),
    );
  }

  void runDesign() {
    if (layers.isEmpty) {
      ScaffoldMessenger.of(context)
          .showSnackBar(const SnackBar(content: Text('Load the well log first')));
      return;
    }

    layers.sort((a, b) => b.kIndex.compareTo(a.kIndex));
    final aquifer = layers.first;
    aquiferTop = aquifer.depthFrom;
    aquiferBottom = aquifer.depthTo;

    final staticWater = double.parse(_staticCtl.text);
    final years = double.parse(_yearsCtl.text);
    final decline = double.parse(_declineCtl.text);
    final drawdown = double.parse(_drawdownCtl.text);
    const safety = 15.0;

    housingDepth = staticWater + decline * years + drawdown + safety;
    screenStart = housingDepth + 10.0;
    final available = aquiferBottom - screenStart;
    if (available < 25) {
      screenLen = 0.7 * available;
    } else if (available <= 50) {
      screenLen = 0.75 * available;
    } else {
      screenLen = 0.8 * available;
    }
    screenLen = screenLen.clamp(0, 100);
    screenEnd = screenStart + screenLen;
    sandTrapLen = 10;

    final scrLayers = layers.where((l) => l.depthFrom <= screenEnd && l.depthTo >= screenStart).toList();
    governing = scrLayers.reduce((a, b) => a.fm <= b.fm ? a : b);

    multFactor = _map(governing.cu, 1.5, 2.5, 4, 6);
    d30Gravel = governing.d30 * multFactor;
    d10Gravel = governing.d10 * multFactor;

    final slotSizeIn = d10Gravel / 25.4;
    rawSlotNo = slotSizeIn * 1000.0;
    const availSlots = [
      3,4,5,6,7,8,10,12,14,16,18,20,25,30,35,40,45,50,60,70,80,100,120,140,170,200,230,270,325,400
    ];
    selectedSlot = availSlots.where((s) => s <= rawSlotNo).fold<double>(availSlots.first.toDouble(), (p, e) => e > p ? e.toDouble() : p);

    openingPct = (selectedSlot / 2) / 100;
    const dStrainerIn = 6.0, criticalVel = 0.1, formLoss = 2.5;
    final qCusec = (openingPct * math.pi * (dStrainerIn/12) * screenLen * criticalVel) / formLoss;
    qLPH = qCusec * math.pow(0.3048, 3) * 3600 * 1000;

    setState(() {});
  }

  double _map(double x, double a0, double a1, double b0, double b1) {
    final t = ((x - a0) / (a1 - a0)).clamp(0, 1);
    return b0 + (b1 - b0) * t;
  }

  Widget _field(String label, TextEditingController controller) {
    return TextField(
      controller: controller,
      decoration: InputDecoration(
        labelText: label,
        border: OutlineInputBorder(borderRadius: BorderRadius.circular(10)),
      ),
      keyboardType: TextInputType.number,
    );
  }

  @override
  Widget build(BuildContext context) {
    final headline = Theme.of(context).textTheme.titleLarge;
    return Scaffold(
      appBar: AppBar(title: const Text('Percola: Tubewell Designer')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          Text('How to use', style: headline),
          const SizedBox(height: 6),
          const Text(
            '1️⃣ Load well_log.xlsx (Depth_from_ft, Depth_to_ft, FM, D60_mm, D10_mm, Cu, optional D30_mm)\n'
            '2️⃣ Enter inputs and press Run Design.',
            style: TextStyle(height: 1.4),
          ),
          const SizedBox(height: 10),
          ElevatedButton.icon(
            onPressed: pickExcel,
            icon: const Icon(Icons.upload_file),
            label: const Text('Load Excel'),
          ),
          const SizedBox(height: 10),
          Row(
            children: [
              Expanded(child: _field('Static water table (ft)', _staticCtl)),
              const SizedBox(width: 10),
              Expanded(child: _field('Design period (yrs)', _yearsCtl)),
            ],
          ),
          const SizedBox(height: 10),
          Row(
            children: [
              Expanded(child: _field('Decline rate (ft/yr)', _declineCtl)),
              const SizedBox(width: 10),
              Expanded(child: _field('Drawdown (ft)', _drawdownCtl)),
            ],
          ),
          const SizedBox(height: 12),
          ElevatedButton.icon(
            onPressed: runDesign,
            icon: const Icon(Icons.calculate),
            label: const Text('Run Design'),
          ),
          if (screenLen > 0) ...[
            const Divider(),
            Text('Results Summary', style: headline),
            const SizedBox(height: 8),
            _resultTile('Housing pipe depth', '${housingDepth.toStringAsFixed(1)} ft'),
            _resultTile('Screen depth range', '${screenStart.toStringAsFixed(1)}–${screenEnd.toStringAsFixed(1)} ft'),
            _resultTile('Screen length', '${screenLen.toStringAsFixed(1)} ft'),
            _resultTile('Sand trap', '${sandTrapLen.toStringAsFixed(1)} ft'),
            _resultTile('Slot number', '#${selectedSlot.toStringAsFixed(0)}'),
            _resultTile('Gravel D10/D30', '${d10Gravel.toStringAsFixed(3)} / ${d30Gravel.toStringAsFixed(3)} mm'),
            _resultTile('Opening area', '${(openingPct * 100).toStringAsFixed(1)} %'),
            _resultTile('Estimated yield', '${qLPH.toStringAsFixed(0)} LPH'),
            const SizedBox(height: 20),
            Text('Gradation Chart', style: headline),
            const SizedBox(height: 8),
            SizedBox(height: 250, child: _buildGradationChart()),
            const SizedBox(height: 20),
            Text('Tentative Layout', style: headline),
            const SizedBox(height: 8),
            SizedBox(height: 400, child: CustomPaint(painter: TubewellPainter(screenStart, screenEnd, sandTrapLen))),
          ],
        ],
      ),
    );
  }

  Widget _resultTile(String title, String value) => Card(
        elevation: 2,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
        child: ListTile(
          title: Text(title),
          trailing: Text(value, style: const TextStyle(fontWeight: FontWeight.bold)),
        ),
      );
  Widget _buildGradationChart() {
    // Define the soil curve (log10 of sieve size)
    final soilPoints = [
      FlSpot(math.log(0.15) / math.ln10, 0),
      FlSpot(math.log(governing.d10) / math.ln10, 10),
      FlSpot(math.log(governing.d30) / math.ln10, 30),
      FlSpot(math.log(governing.d60) / math.ln10, 60),
      FlSpot(math.log(4.75) / math.ln10, 100),
    ];

    final gravelK = multFactor;
    final gravelPoints = soilPoints
        .map((p) => FlSpot(p.x + math.log(gravelK) / math.ln10, p.y))
        .toList();

    // X-axis tick positions (log10 scale) for 0.01, 0.1, 1, 10, 100 mm

    return LineChart(
      LineChartData(
        minX: -2,
        maxX: 2,
        minY: 0,
        maxY: 100,
        gridData: const FlGridData(show: true),
        borderData: FlBorderData(show: true),
        titlesData: FlTitlesData(
          leftTitles: AxisTitles(
            sideTitles: SideTitles(
              reservedSize: 38,
              showTitles: true,
              getTitlesWidget: (y, meta) =>
                  Text('${y.toStringAsFixed(0)}%', style: const TextStyle(fontSize: 11)),
              interval: 20,
            ),
          ),
          bottomTitles: AxisTitles(
            axisNameWidget: const Text('Particle size (mm)', style: TextStyle(fontWeight: FontWeight.bold)),
            sideTitles: SideTitles(
              reservedSize: 32,
              showTitles: true,
              getTitlesWidget: (x, meta) {
                final sieve = math.pow(10, x).toStringAsFixed(2);
                return Text(sieve, style: const TextStyle(fontSize: 11));
              },
              interval: 1,
            ),
          ),
          topTitles: const AxisTitles(sideTitles: SideTitles(showTitles: false)),
          rightTitles: const AxisTitles(sideTitles: SideTitles(showTitles: false)),
        ),
        lineBarsData: [
          LineChartBarData(
            spots: soilPoints,
            color: Colors.blue.shade700,
            isCurved: true,
            barWidth: 3,
            dotData: const FlDotData(show: true),
          ),
          LineChartBarData(
            spots: gravelPoints,
            color: Colors.green.shade700,
            isCurved: true,
            barWidth: 3,
            dashArray: [6, 3],
            dotData: const FlDotData(show: true),
          ),
        ],
      ),
    );
  }


}
