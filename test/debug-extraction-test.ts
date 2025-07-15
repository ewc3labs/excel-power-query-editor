import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';

/**
 * Manual test script to run debug extractions and analyze results
 */
async function runDebugExtractionTests() {
    console.log('🧪 Running debug extraction tests...');
    
    const testFixturesDir = path.join(__dirname, 'fixtures');
    
    // Test files to extract
    const testFiles = [
        'simple.xlsx',
        'complex.xlsm'
    ];
    
    for (const testFile of testFiles) {
        const filePath = path.join(testFixturesDir, testFile);
        if (!fs.existsSync(filePath)) {
            console.log(`❌ Test file not found: ${testFile}`);
            continue;
        }
        
        console.log(`\n📁 Testing debug extraction: ${testFile}`);
        
        try {
            // Run debug extraction command
            const uri = vscode.Uri.file(filePath);
            await vscode.commands.executeCommand('excel-power-query-editor.rawExtraction', uri);
            
            // Wait for extraction to complete
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Check results
            const baseName = path.basename(testFile, path.extname(testFile));
            const debugDir = path.join(testFixturesDir, `${baseName}_debug_extraction`);
            
            if (fs.existsSync(debugDir)) {
                console.log(`✅ Debug directory created: ${path.basename(debugDir)}`);
                
                // List all files in debug directory
                const files = fs.readdirSync(debugDir, { recursive: true }) as string[];
                console.log(`📊 Generated ${files.length} files:`);
                
                // Categorize files
                const categories = {
                    powerQuery: files.filter(f => f.endsWith('_PowerQuery.m')),
                    reports: files.filter(f => f.includes('REPORT.json')),
                    dataMashup: files.filter(f => f.includes('DATAMASHUP_')),
                    xmlFiles: files.filter(f => f.endsWith('.xml') || f.endsWith('.xml.txt')),
                    other: files.filter(f => !f.endsWith('_PowerQuery.m') && !f.includes('REPORT.json') && !f.includes('DATAMASHUP_') && !f.endsWith('.xml') && !f.endsWith('.xml.txt'))
                };
                
                console.log(`  💾 Power Query M files: ${categories.powerQuery.length}`);
                console.log(`  📋 Report files: ${categories.reports.length}`);
                console.log(`  🔍 DataMashup files: ${categories.dataMashup.length}`);
                console.log(`  📄 XML files: ${categories.xmlFiles.length}`);
                console.log(`  📂 Other files: ${categories.other.length}`);
                
                // Check for main report
                const reportFile = path.join(debugDir, 'EXTRACTION_REPORT.json');
                if (fs.existsSync(reportFile)) {
                    const report = JSON.parse(fs.readFileSync(reportFile, 'utf8'));
                    console.log(`📈 DataMashup files found: ${report.dataMashupAnalysis.dataMashupFilesFound}`);
                    console.log(`📊 Total XML files scanned: ${report.dataMashupAnalysis.totalXmlFilesScanned}`);
                }
                
            } else {
                console.log(`❌ Debug directory not created: ${debugDir}`);
            }
            
        } catch (error) {
            console.log(`❌ Debug extraction failed for ${testFile}: ${error}`);
        }
    }
    
    console.log('\n🏁 Debug extraction tests completed');
}

// Export for use in other tests
export { runDebugExtractionTests };
