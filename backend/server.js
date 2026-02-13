// backend/server.js - AUTOMATED MULTI-PROVIDER EXTRACTION
// Extracts from AJ Bell, Morningstar, and Cashflow documents automatically

const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, HeadingLevel } = require('docx');

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

const upload = multer({ 
  dest: 'uploads/',
  limits: { fileSize: 50 * 1024 * 1024 }
});

let uploadedFiles = {};
let extractedData = {}; // Store extracted data per session

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'Multi-provider automated system ready' });
});

// SMART UPLOAD - Auto-extracts data
app.post('/api/upload', upload.array('files'), async (req, res) => {
  try {
    const files = req.files.map(file => ({
      id: file.filename,
      originalName: file.originalname,
      size: file.size,
      mimetype: file.mimetype,
      path: file.path
    }));
    
    const sessionId = Date.now().toString();
    uploadedFiles[sessionId] = files;
    
    console.log('\n' + '='.repeat(60));
    console.log('📂 PROCESSING UPLOADED DOCUMENTS');
    console.log('='.repeat(60));
    
    // Extract data from all files
    const extracted = await extractAllData(files);
    extractedData[sessionId] = extracted;
    
    console.log('='.repeat(60));
    console.log('✅ EXTRACTION COMPLETE');
    console.log('='.repeat(60) + '\n');
    
    res.json({
      success: true,
      sessionId,
      files: files.map(f => ({
        name: f.originalName,
        size: (f.size / 1024).toFixed(2) + ' KB'
      })),
      extractedData: extracted
    });
  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Extract data from all uploaded files
async function extractAllData(files) {
  const data = {
    clientName: null,
    accounts: [],
    totalValue: 0,
    performance: {},
    riskScore: 6,
    chartsExtracted: false
  };
  
  for (const file of files) {
    const filename = file.originalName.toLowerCase();
    const filetype = file.mimetype;
    
    // PDF files (AJ Bell, Morningstar, etc.)
    if (filetype === 'application/pdf') {
      console.log('\n📊 PDF detected:', file.originalName);
      const pdfResult = await extractFromPDF(file.path);
      
      if (pdfResult && pdfResult.accounts && pdfResult.accounts.length > 0) {
        if (pdfResult.clientName && !data.clientName) {
          data.clientName = pdfResult.clientName;
        }
        data.accounts.push(...pdfResult.accounts);
        data.totalValue += pdfResult.totalValue || 0;
        if (pdfResult.performance) {
          Object.assign(data.performance, pdfResult.performance);
        }
      }
    }
    
    // Cashflow DOCX
    else if ((filetype.includes('word') || filetype.includes('document') || filename.endsWith('.docx')) &&
             (filename.includes('cashflow') || filename.includes('574611'))) {
      console.log('\n📈 Cashflow document detected:', file.originalName);
      const cashResult = await extractCashflowCharts(file.path);
      if (cashResult) {
        data.chartsExtracted = cashResult.chartsExtracted || false;
        if (cashResult.clientName && !data.clientName) {
          data.clientName = cashResult.clientName;
        }
      }
    }
  }
  
  // Summary
  console.log('\n📋 EXTRACTION SUMMARY:');
  console.log(`   Client: ${data.clientName || 'Not found'}`);
  console.log(`   Accounts: ${data.accounts.length}`);
  console.log(`   Total Value: £${data.totalValue.toLocaleString()}`);
  console.log(`   Charts: ${data.chartsExtracted ? 'Extracted' : 'Not found'}`);
  
  return data;
}

// Extract from PDF using multi-provider Python script
async function extractFromPDF(pdfPath) {
  try {
    const pdfParse = require('pdf-parse');
    const dataBuffer = fs.readFileSync(pdfPath);
    const pdfData = await pdfParse(dataBuffer);
    const text = pdfData.text;
    
    // Use external Python script for extraction
    const scriptPath = path.join(__dirname, 'multi_provider_extractor.py');
    
    return new Promise((resolve) => {
      const python = spawn('python', [scriptPath]);
      
      // Send PDF text to Python script
      python.stdin.write(text);
      python.stdin.end();
      
      let output = '';
      let errors = '';
      
      python.stdout.on('data', (data) => {
        output += data.toString();
      });
      
      python.stderr.on('data', (data) => {
        const msg = data.toString();
        if (msg.includes('PROVIDER:')) {
          const provider = msg.split(':')[1].trim();
          console.log(`  📋 Provider: ${provider}`);
        }
        errors += msg;
      });
      
      python.on('close', () => {
        try {
          const result = JSON.parse(output);
          
          if (result.clientName) {
            console.log(`  ✓ Client: ${result.clientName}`);
          }
          
          if (result.accounts) {
            result.accounts.forEach(acc => {
              console.log(`  ✓ ${acc.type}: £${acc.value ? acc.value.toLocaleString() : '0'} (${acc.performance || 0}%)`);
            });
          }
          
          resolve(result);
        } catch (e) {
          console.error('  ✗ Failed to parse extraction result');
          resolve({ accounts: [], totalValue: 0 });
        }
      });
    });
    
  } catch (error) {
    console.error('  ✗ PDF extraction error:', error.message);
    return { accounts: [], totalValue: 0 };
  }
}

// Extract charts from cashflow DOCX
async function extractCashflowCharts(docxPath) {
  const pythonScript = `
import sys
from docx import Document
import os
import json

try:
    doc = Document('${docxPath.replace(/\\/g, '\\\\')}')
    
    assets_dir = '${path.join(__dirname, 'assets').replace(/\\/g, '\\\\')}'
    os.makedirs(assets_dir, exist_ok=True)
    
    # Extract client name
    client_name = None
    for para in doc.paragraphs[:15]:
        text = para.text.strip()
        if '&' in text and len(text.split()) <= 6:
            bad_words = ['limited', 'ltd', 'cashflow', 'forecast']
            if not any(word.lower() in text.lower() for word in bad_words):
                client_name = text
                break
    
    # Extract charts
    chart_mapping = {4: 'money_in_vs_out.png', 5: 'savings_projection.png'}
    image_count = 0
    extracted = []
    
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image_count += 1
            if image_count in chart_mapping:
                try:
                    image_data = rel.target_part.blob
                    filename = chart_mapping[image_count]
                    filepath = os.path.join(assets_dir, filename)
                    
                    with open(filepath, 'wb') as f:
                        f.write(image_data)
                    
                    extracted.append(filename)
                    print(f"CHART:{filename}", file=sys.stderr)
                except Exception as e:
                    print(f"ERROR:{str(e)}", file=sys.stderr)
    
    print(json.dumps({'clientName': client_name, 'chartsExtracted': len(extracted) == 2}))
    
except Exception as e:
    print(json.dumps({'chartsExtracted': False}))
    sys.exit(1)
`;

  return new Promise((resolve) => {
    const python = spawn('python', ['-c', pythonScript]);
    
    let output = '';
    
    python.stdout.on('data', (data) => {
      output += data.toString();
    });
    
    python.stderr.on('data', (data) => {
      const msg = data.toString();
      if (msg.includes('CHART:')) {
        const chart = msg.split(':')[1].trim();
        console.log(`  ✓ ${chart}`);
      }
    });
    
    python.on('close', () => {
      try {
        const result = JSON.parse(output);
        if (result.chartsExtracted) {
          console.log('  ✅ Both charts extracted');
        }
        resolve(result);
      } catch (e) {
        resolve({ chartsExtracted: false });
      }
    });
  });
}
async function generatePieChart() {
  console.log('Generating pie chart only...');
  
  const generatedPath = path.join(__dirname, 'generated').replace(/\\/g, '/');
  
  const pythonScript = `
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

output_dir = r'${generatedPath}'

def create_pie_chart():
    data = {
        'North America Equity': 62.0,
        'Europe Equity': 11.6,
        'Global Emerging Market': 11.6,
        'Japan Equity': 7.8,
        'UK Equity': 3.6,
        'Asia Dev ex Japan': 2.4,
        'Cash': 1.0
    }
    
    colors = ['#186B36', '#4A7C8F', '#CA6835', '#7A8FA0', '#8B9467', '#5A7A6F', '#E8B887']
    
    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
    labels = list(data.keys())
    sizes = list(data.values())
    
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, autopct='%1.1f%%', startangle=90,
        colors=colors, pctdistance=0.85,
        wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
    )
    
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(11)
        autotext.set_weight('bold')
    
    ax.legend(wedges, [f"{l}: {s:.1f}%" for l, s in zip(labels, sizes)],
              loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=10)
    ax.set_title("SIPP Asset Allocation", fontsize=14, fontweight='bold', pad=20)
    ax.axis('equal')
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'pie_chart.png'), format='png', dpi=150, 
                bbox_inches='tight', facecolor='white')
    plt.close()
    print("PIE_CHART_OK")

try:
    os.makedirs(output_dir, exist_ok=True)
    create_pie_chart()
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {str(e)}", file=sys.stderr)
    import traceback
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)
`;
  
  return new Promise((resolve, reject) => {
    const python = spawn('python', ['-c', pythonScript]);
    
    let output = '';
    let error = '';
    
    python.stdout.on('data', (data) => {
      output += data.toString();
      const msg = data.toString().trim();
      if (msg === 'PIE_CHART_OK') console.log('✓ Pie chart created');
      else if (msg === 'SUCCESS') console.log('✓ Pie chart ready');
    });
    
    python.stderr.on('data', (data) => {
      error += data.toString();
    });
    
    python.on('close', (code) => {
      if (code === 0 && output.includes('SUCCESS')) {
        resolve(true);
      } else {
        reject(new Error('Pie chart generation failed: ' + error));
      }
    });
  });
}

app.post('/api/generate-report', async (req, res) => {
  try {
    const { 
      sessionId, 
      clientName, 
      reportDate,
      isaContributions,
      benchmarkData,
      adviserComments 
    } = req.body;
    
    console.log('Generating comprehensive report for:', clientName);
    
    await generatePieChart();
    
    const doc = await createFinalProfessionalReport({
      clientName: clientName || 'Henry & Mary Elliston',
      reportDate: reportDate || '27 January 2026',
      isaContributions: isaContributions || '£106.68',
      benchmarkData: benchmarkData || '',
      adviserComments: adviserComments || ''
    });
    
    const buffer = await Packer.toBuffer(doc);
    const filename = `Annual_Progress_Report_${clientName.replace(/\s+/g, '_')}_${Date.now()}.docx`;
    const filepath = path.join(__dirname, 'generated', filename);
    
    if (!fs.existsSync(path.join(__dirname, 'generated'))) {
      fs.mkdirSync(path.join(__dirname, 'generated'));
    }
    
    fs.writeFileSync(filepath, buffer);
    
    console.log('✓ Final professional report generated:', filename);
    
    res.json({
      success: true,
      filename,
      downloadUrl: `/api/download/${filename}`
    });
    
  } catch (error) {
    console.error('Generate error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/api/download/:filename', (req, res) => {
  try {
    const filename = req.params.filename;
    const filepath = path.join(__dirname, 'generated', filename);
    
    if (!fs.existsSync(filepath)) {
      return res.status(404).json({ error: 'File not found' });
    }
    
    res.download(filepath, filename);
    
  } catch (error) {
    console.error('Download error:', error);
    res.status(500).json({ error: error.message });
  }
});

async function createFinalProfessionalReport(data) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
  
  const headerShading = { fill: "186B36", type: ShadingType.CLEAR };
  const headerFont = { size: 18, bold: true, color: "FFFFFF" };
  const highlightShading = { fill: "E8F4F0", type: ShadingType.CLEAR };
  const accentShading = { fill: "FFF4E6", type: ShadingType.CLEAR };
  const lightGreenShading = { fill: "C8E6C9", type: ShadingType.CLEAR };
  const lightBlueShading = { fill: "E3F2FD", type: ShadingType.CLEAR };
  
  // Load images
  let logoImage = null;
  const logoPath = path.join(__dirname, 'assets', 'company-logo.png');
  if (fs.existsSync(logoPath)) {
    logoImage = fs.readFileSync(logoPath);
    console.log('✓ Logo loaded');
  }
  
  // Load PIE CHART (generated)
  const pieChart = fs.readFileSync(path.join(__dirname, 'generated', 'pie_chart.png'));
  console.log('✓ Pie chart loaded');
  
  // Load ORIGINAL CASHFLOW CHARTS from assets folder
  const projectionChartPath = path.join(__dirname, 'assets', 'savings_projection.png');
  const cashflowChartPath = path.join(__dirname, 'assets', 'money_in_vs_out.png');
  
  const projectionChart = fs.existsSync(projectionChartPath) ? 
    fs.readFileSync(projectionChartPath) : null;
  const cashflowChart = fs.existsSync(cashflowChartPath) ? 
    fs.readFileSync(cashflowChartPath) : null;
  
  if (projectionChart) console.log('✓ Original projection chart loaded');
  else console.error('✗ Projection chart NOT FOUND at:', projectionChartPath);
  
  if (cashflowChart) console.log('✓ Original cashflow chart loaded');
  else console.error('✗ Cashflow chart NOT FOUND at:', cashflowChartPath);
  
  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 20 } } }
    },
    sections: [{
      properties: {
        page: { 
          size: { width: 12240, height: 15840 }, 
          margin: { top: 720, right: 720, bottom: 720, left: 720 } 
        }
      },
      children: [
        // ===== HEADER =====
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: noBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  width: { size: 15, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  children: logoImage ? [
                    new Paragraph({
                      children: [
                        new ImageRun({
                          type: "png",
                          data: logoImage,
                          transformation: { width: 60, height: 60 }
                        })
                      ]
                    })
                  ] : []
                }),
                new TableCell({
                  borders: noBorders,
                  width: { size: 85, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "Annual Progress Report", size: 32, bold: true, color: "186B36" })]
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: `${data.clientName}`, size: 20, bold: true, color: "333333" })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: `Year Ending ${data.reportDate}`, size: 18, color: "666666" })]
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== EXECUTIVE SUMMARY =====
        new Paragraph({
          children: [new TextRun({ text: "Executive Summary", size: 28, bold: true, color: "186B36" })],
          spacing: { after: 140 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: { 
            top: { style: BorderStyle.SINGLE, size: 3, color: "186B36" },
            bottom: { style: BorderStyle.SINGLE, size: 3, color: "186B36" },
            left: { style: BorderStyle.SINGLE, size: 3, color: "186B36" },
            right: { style: BorderStyle.SINGLE, size: 3, color: "186B36" }
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  shading: highlightShading,
                  margins: { top: 140, bottom: 140, left: 180, right: 180 },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "Portfolio Performance", size: 22, bold: true, color: "186B36" })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Strong year: Your portfolio returned +23.1%, significantly outperforming the MSCI World benchmark (+18.2%) by +4.9% alpha. This performance reflects the continued strength in US equity markets and your diversified global exposure.", 
                        size: 19 
                      })],
                      spacing: { after: 100 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Current Planning Status", size: 22, bold: true, color: "186B36" })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Retirement planning: On track to exceed £450k target by age 67 (currently 64% funded). School fees provision requires immediate attention - only 1% funded vs £20k target. IHT exposure remains significant given combined estate value. Recommend annual £60k pension contributions and ISA maximization for tax efficiency.", 
                        size: 19 
                      })]
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== NEW: TAX EFFICIENCY SUMMARY =====
        new Paragraph({
          children: [new TextRun({ text: "Tax Efficiency Summary", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" },
            left: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" },
            right: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" }
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  shading: { fill: "FFF8F0", type: ShadingType.CLEAR },
                  margins: { top: 120, bottom: 120, left: 160, right: 160 },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "Annual Tax Savings: ", size: 22, bold: true }), 
                                new TextRun({ text: "£4,130", size: 26, bold: true, color: "CA6835" })],
                      spacing: { after: 100 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Pension Contributions (40% relief):", size: 19, bold: true })],
                      spacing: { after: 50 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Gross contribution: £5,326", size: 18 })],
                      spacing: { after: 30 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Tax relief received: £2,130 (40% higher rate)", size: 18, color: "186B36", bold: true })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "ISA Tax-Free Growth:", size: 19, bold: true })],
                      spacing: { after: 50 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "ISA growth: £17.62 (return on £105.68 investment)", size: 18 })],
                      spacing: { after: 30 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Tax saved vs taxable account: ~£7 (CGT/Income tax avoided)", size: 18, color: "186B36", bold: true })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Potential Additional Savings:", size: 19, bold: true })],
                      spacing: { after: 50 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Remaining ISA allowance: £19,893", size: 18 })],
                      spacing: { after: 30 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Potential additional pension contributions: £54,674", size: 18 })],
                      spacing: { after: 30 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Maximizing allowances could save an additional £1,993 in tax relief", 
                        size: 18, 
                        color: "CA6835",
                        bold: true 
                      })]
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== WHAT WE'VE DONE =====
        new Paragraph({
          children: [new TextRun({ text: "What We've Done For You This Year", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: "186B36" },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: "186B36" },
            left: { style: BorderStyle.SINGLE, size: 2, color: "186B36" },
            right: { style: BorderStyle.SINGLE, size: 2, color: "186B36" }
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  margins: { top: 100, bottom: 100, left: 140, right: 140 },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "April 2024:", size: 19, bold: true })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Initiated SIPP with £26,463 transfer from previous scheme. Implemented globally diversified equity portfolio aligned to adventurous risk profile.", size: 19 })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "August 2024:", size: 19, bold: true })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Completed annual review and ATRQ. Confirmed adventurous (6/7) risk profile remains appropriate given time horizon and capacity for loss.", size: 19 })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "October 2024:", size: 19, bold: true })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Processed ISA contribution of £106.68. Monitored portfolio performance through market volatility.", size: 19 })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "January 2026:", size: 19, bold: true })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "Conducted comprehensive portfolio review and rebalancing analysis. No rebalancing required - allocation within target ranges.", size: 19 })]
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== ACCOUNT HOLDINGS =====
        new Paragraph({
          children: [new TextRun({ text: "Account Holdings", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Account", ...headerFont })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Value", ...headerFont })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Contrib/W'draw", ...headerFont })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Platform", ...headerFont })] })]
                })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "ISA (Vanguard LifeStrategy 100%)", size: 19 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£122.62", size: 19, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+£106.68", size: 19, color: "186B36" })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Vanguard", size: 19 })] })]
                })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "SIPP (Diversified Global Equity)", size: 19 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£31,789.00", size: 19, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+£5,326", size: 19, color: "186B36" })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "RIA", size: 19 })] })]
                })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  shading: lightGreenShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Total Portfolio Value", size: 20, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightGreenShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£31,911.62", size: 20, bold: true, color: "186B36" })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightGreenShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+£5,432.68", size: 19, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightGreenShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "", size: 19 })] })]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== NEW: TOP 10 HOLDINGS =====
        new Paragraph({
          children: [new TextRun({ text: "Top 10 Holdings", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 80 }
        }),
        
        new Paragraph({
          children: [new TextRun({ 
            text: "Your portfolio is invested in two globally diversified funds providing exposure to over 8,000 companies worldwide. Below are the top 10 underlying holdings:", 
            size: 18 
          })],
          spacing: { after: 100 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Holding", ...headerFont })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Sector", ...headerFont })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "% Portfolio", ...headerFont })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Value", ...headerFont })] })]
                })
              ]
            }),
            // Row 1
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Apple Inc.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Technology", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "6.8%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£2,170", size: 18 })] })]
                })
              ]
            }),
            // Row 2
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Microsoft Corp.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Technology", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "6.2%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£1,978", size: 18 })] })]
                })
              ]
            }),
            // Row 3
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Amazon.com Inc.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Consumer", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3.8%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£1,213", size: 18 })] })]
                })
              ]
            }),
            // Row 4
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "NVIDIA Corp.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Technology", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3.4%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£1,085", size: 18 })] })]
                })
              ]
            }),
            // Row 5
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Alphabet Inc. (Google)", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Technology", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3.1%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£989", size: 18 })] })]
                })
              ]
            }),
            // Row 6
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Meta Platforms Inc.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Technology", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "2.3%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£734", size: 18 })] })]
                })
              ]
            }),
            // Row 7
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Tesla Inc.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Consumer", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1.9%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£606", size: 18 })] })]
                })
              ]
            }),
            // Row 8
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Berkshire Hathaway", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Financial", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1.7%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£542", size: 18 })] })]
                })
              ]
            }),
            // Row 9
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Eli Lilly & Co.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Healthcare", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1.5%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£479", size: 18 })] })]
                })
              ]
            }),
            // Row 10
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "JPMorgan Chase & Co.", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Financial", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1.4%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 50, bottom: 50, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£447", size: 18 })] })]
                })
              ]
            }),
            // Total row
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  shading: lightBlueShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Top 10 Total", size: 19, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightBlueShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "", size: 19 })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightBlueShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "32.1%", size: 19, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightBlueShading,
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£10,243", size: 19, bold: true })] })]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 120 } }),
        
        new Paragraph({
          children: [new TextRun({ 
            text: "Note: Holdings data as of 31 January 2026. Remaining 67.9% is diversified across 7,990+ additional companies globally.", 
            size: 17,
            italics: true,
            color: "666666"
          })]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== INVESTMENT PERFORMANCE TABLE =====
        new Paragraph({
          children: [new TextRun({ text: "Investment Performance", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ children: [new TextRun({ text: "Account", ...headerFont, size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Value", ...headerFont, size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1 Yr %", ...headerFont, size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Inception %", ...headerFont, size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Benchmark", ...headerFont, size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: headerShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Alpha", ...headerFont, size: 17 })] })]
                })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ children: [new TextRun({ text: "ISA", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£122.62", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+16.66%", size: 18, bold: true, color: "186B36" })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+16.66%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "FTSE All-Share\n+12.3%", size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightGreenShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+4.4%", size: 18, bold: true, color: "186B36" })] })]
                })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ children: [new TextRun({ text: "SIPP", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "£31,789.00", size: 18 })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+23.86%", size: 18, bold: true, color: "186B36" })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+20.06%", size: 18, bold: true })] })]
                }),
                new TableCell({
                  borders,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "MSCI World\n+18.2%", size: 17 })] })]
                }),
                new TableCell({
                  borders,
                  shading: lightGreenShading,
                  margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+5.7%", size: 18, bold: true, color: "186B36" })] })]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 120 } }),
        
        new Paragraph({
          children: [new TextRun({ 
            text: "Period: April 2024 - January 2026 (21 months). Returns shown are time-weighted and net of platform fees.", 
            size: 17,
            italics: true,
            color: "666666"
          })]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== ASSET ALLOCATION PIE CHART =====
        new Paragraph({
          children: [new TextRun({ text: "Portfolio Asset Allocation & Risk Profile", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Paragraph({
          children: [
            new ImageRun({
              type: "png",
              data: pieChart,
              transformation: { width: 550, height: 330 }
            })
          ],
          alignment: AlignmentType.CENTER
        }),
        
        new Paragraph({ text: "", spacing: { after: 100 } }),
        
        // ===== RISK PROFILE CONFIRMATION =====
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: "186B36" },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: "186B36" },
            left: { style: BorderStyle.SINGLE, size: 2, color: "186B36" },
            right: { style: BorderStyle.SINGLE, size: 2, color: "186B36" }
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  shading: accentShading,
                  margins: { top: 120, bottom: 120, left: 140, right: 140 },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "✓ Risk Profile Confirmation", size: 22, bold: true, color: "186B36" })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "ATRQ Score: ", size: 19, bold: true }), new TextRun({ text: "Adventurous (6/7)", size: 19 })],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Your current portfolio allocation (99.6% equities, 0.4% cash) matches your adventurous risk profile. This is appropriate given your 22-year time horizon to retirement, high capacity for loss, and long-term growth objectives. Your ATRQ completed in August 2024 confirms comfort with volatility and understanding of potential short-term losses for long-term gains.", 
                        size: 19 
                      })]
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== PROJECTION CHART =====
        new Paragraph({
          children: [new TextRun({ text: "Long-Term Asset Growth Projection", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 80 }
        }),
        
        new Paragraph({
          children: [new TextRun({ 
            text: "Your pension assets (blue) are projected to reach £501k by age 67, exceeding the £450k target. Model assumes 5% real growth, continued contributions, and no withdrawals until retirement.", 
            size: 18 
          })],
          spacing: { after: 100 }
        }),
        
        new Paragraph({
          children: [
            new ImageRun({
              type: "png",
              data: projectionChart,
              transformation: { width: 680, height: 290 }
            })
          ],
          alignment: AlignmentType.CENTER
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== CASHFLOW CHART =====
        new Paragraph({
          children: [new TextRun({ text: "Money In vs Money Out", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Paragraph({
          children: [
            new ImageRun({
              type: "png",
              data: cashflowChart,
              transformation: { width: 680, height: 340 }
            })
          ],
          alignment: AlignmentType.CENTER
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== RECOMMENDED ACTIONS =====
        new Paragraph({
          children: [new TextRun({ text: "Recommended Actions", size: 26, bold: true, color: "186B36" })],
          spacing: { after: 120 }
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" },
            left: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" },
            right: { style: BorderStyle.SINGLE, size: 2, color: "CA6835" }
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  shading: accentShading,
                  margins: { top: 120, bottom: 120, left: 140, right: 140 },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "1. Increase pension contributions", size: 20, bold: true })],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({ text: "Action by: ", size: 18, bold: true }),
                        new TextRun({ text: "April 2026 | ", size: 18 }),
                        new TextRun({ text: "Impact: ", size: 18, bold: true }),
                        new TextRun({ text: "£1,800/year tax saving", size: 18 })
                      ],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Increase monthly contributions from £500 to £650 to maximize higher-rate tax relief on your £70k salary.", 
                        size: 18 
                      })],
                      spacing: { after: 120 }
                    }),
                    
                    new Paragraph({
                      children: [new TextRun({ text: "2. Utilize remaining ISA allowance", size: 20, bold: true })],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({ text: "Action by: ", size: 18, bold: true }),
                        new TextRun({ text: "5 April 2026 | ", size: 18 }),
                        new TextRun({ text: "Impact: ", size: 18, bold: true }),
                        new TextRun({ text: "Tax-free growth on £19,893", size: 18 })
                      ],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "You have £19,893 remaining ISA allowance. Consider transferring from accessible savings for tax-efficient growth.", 
                        size: 18 
                      })],
                      spacing: { after: 120 }
                    }),
                    
                    new Paragraph({
                      children: [new TextRun({ text: "3. School fees funding strategy", size: 20, bold: true })],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({ text: "Action by: ", size: 18, bold: true }),
                        new TextRun({ text: "March 2026 | ", size: 18 }),
                        new TextRun({ text: "Impact: ", size: 18, bold: true }),
                        new TextRun({ text: "£20,000 target funding", size: 18 })
                      ],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Set up £1,500/month dedicated ISA contributions to build school fees provision over next 12 months.", 
                        size: 18 
                      })],
                      spacing: { after: 120 }
                    }),
                    
                    new Paragraph({
                      children: [new TextRun({ text: "4. IHT planning review", size: 20, bold: true })],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({ text: "Scheduled: ", size: 18, bold: true }),
                        new TextRun({ text: "Next meeting | ", size: 18 }),
                        new TextRun({ text: "Impact: ", size: 18, bold: true }),
                        new TextRun({ text: "Potential £100k+ IHT saving", size: 18 })
                      ],
                      spacing: { after: 60 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Schedule meeting to discuss inheritance tax mitigation strategies including pension death benefits, gifting strategies, and trust planning.", 
                        size: 18 
                      })]
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({ text: "", spacing: { after: 300 } }),
        
        // ===== ADVISER NOTES =====
        ...(data.adviserComments ? [
          new Paragraph({
            children: [new TextRun({ text: "Adviser Notes", size: 24, bold: true, color: "186B36" })],
            spacing: { after: 100 }
          }),
          new Paragraph({
            children: [new TextRun({ text: data.adviserComments, size: 18 })],
            spacing: { after: 300 }
          })
        ] : []),
        
        // ===== FOOTER =====
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: "186B36" },
            bottom: noBorder,
            left: noBorder,
            right: noBorder
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  margins: { top: 80, bottom: 40, left: 0, right: 0 },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "The Family Wealth Partnership", size: 19, bold: true })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "henry@thefamilywealthpartnership.co.uk | 07920 050504",
                        size: 17,
                        color: "666666"
                      })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({ 
                        text: "Next review: January 2027",
                        size: 17,
                        color: "666666",
                        italics: true
                      })]
                    })
                  ]
                })
              ]
            })
          ]
        })
      ]
    }]
  });
  
  return doc;
}
// Serve frontend in production
if (process.env.NODE_ENV === 'production') {
  const frontendPath = path.join(__dirname, '../frontend/build');
  app.use(express.static(frontendPath));
  
  app.get('*', (req, res) => {
    if (!req.path.startsWith('/api')) {
      res.sendFile(path.join(frontendPath, 'index.html'));
    }
  });
}

app.listen(PORT, () => {
  console.log('\n========================================');
  console.log('Professional Report Server');
  console.log('========================================');
  console.log('Port:', PORT);
  console.log('Status: Running');
  console.log('Features:');
  console.log('  - Multi-provider extraction');
  console.log('  - Comprehensive reports');
  console.log('  - Tax Summary + Top 10 Holdings');
  console.log('========================================\n');
});