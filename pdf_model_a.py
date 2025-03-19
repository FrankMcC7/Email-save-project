import os
import json
import pickle
import argparse
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional, Union

import pandas as pd
from PyPDF2 import PdfReader
from tqdm import tqdm
import logging
import hashlib

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger('adaptive_pdf_extractor')

class ExtractionRule:
    """Represents a rule for extracting specific data from PDFs."""
    
    def __init__(self, 
                 name: str, 
                 pattern: str, 
                 description: str = None,
                 match_index: int = 0,
                 match_group: int = 0,
                 preprocessing: List[str] = None):
        """
        Initialize an extraction rule.
        
        Args:
            name: Unique identifier for this extraction rule
            pattern: Regex pattern to extract the value
            description: Human-readable description of what this rule extracts
            match_index: Which match to use if multiple are found (default: first)
            match_group: Which capture group to use from the match (default: 0 = entire match)
            preprocessing: List of preprocessing functions to apply to text before extraction
        """
        self.name = name
        self.pattern = pattern
        self.description = description or f"Extract {name}"
        self.match_index = match_index
        self.match_group = match_group
        self.preprocessing = preprocessing or []
        self.compiled_pattern = re.compile(pattern, re.DOTALL | re.MULTILINE)
        self.corrections = {}  # Map of document fingerprints to corrected values
        
    def extract(self, text: str, doc_id: str = None) -> Tuple[Optional[str], bool]:
        """
        Extract the value from text.
        
        Args:
            text: Text to extract from
            doc_id: Document identifier for checking corrections
            
        Returns:
            Tuple of (extracted_value, is_correction)
        """
        # Check if we have a correction for this specific document
        if doc_id and doc_id in self.corrections:
            return self.corrections[doc_id], True
        
        # Apply preprocessing if any
        processed_text = text
        for process in self.preprocessing:
            if process == "lowercase":
                processed_text = processed_text.lower()
            elif process == "remove_whitespace":
                processed_text = re.sub(r'\s+', '', processed_text)
                
        # Try to match the pattern
        matches = self.compiled_pattern.findall(processed_text)
        
        if not matches:
            return None, False
            
        if isinstance(matches[0], tuple) and self.match_group < len(matches[0]):
            # If we have capture groups
            try:
                value = matches[min(self.match_index, len(matches)-1)][self.match_group]
                return value.strip(), False
            except (IndexError, TypeError):
                return None, False
        else:
            # If we don't have capture groups or want the entire match
            try:
                value = matches[min(self.match_index, len(matches)-1)]
                return value.strip() if isinstance(value, str) else value, False
            except (IndexError, TypeError):
                return None, False
                
    def add_correction(self, doc_id: str, correct_value: str) -> None:
        """Add a correction for a specific document."""
        self.corrections[doc_id] = correct_value
        logger.info(f"Added correction for rule '{self.name}' on document '{doc_id}': {correct_value}")


class PDFKnowledgeBase:
    """Knowledge base for PDF extraction with learning capabilities."""
    
    def __init__(self, kb_path: str = "pdf_knowledge_base.pkl"):
        """
        Initialize the knowledge base.
        
        Args:
            kb_path: Path to save/load the knowledge base
        """
        self.kb_path = kb_path
        self.rules = {}  # Map of rule_name -> ExtractionRule
        self.document_fingerprints = {}  # Map of filename -> doc_id
        self.load_kb()
        
    def add_rule(self, rule: ExtractionRule) -> None:
        """Add a new extraction rule."""
        self.rules[rule.name] = rule
        logger.info(f"Added extraction rule: {rule.name}")
        
    def get_rule(self, rule_name: str) -> Optional[ExtractionRule]:
        """Get a rule by name."""
        return self.rules.get(rule_name)
        
    def get_document_id(self, filepath: str, content_hash: str = None) -> str:
        """
        Get a unique ID for a document, creating one if it doesn't exist.
        Uses both filename and content hash to identify documents.
        """
        filename = os.path.basename(filepath)
        if filename in self.document_fingerprints:
            return self.document_fingerprints[filename]
            
        # Create a new fingerprint
        if not content_hash:
            # Calculate hash if not provided
            with open(filepath, 'rb') as f:
                content = f.read()
                content_hash = hashlib.md5(content).hexdigest()
                
        doc_id = f"{filename}_{content_hash[:8]}"
        self.document_fingerprints[filename] = doc_id
        return doc_id
        
    def add_correction(self, doc_id: str, rule_name: str, correct_value: str) -> None:
        """Add a correction for a specific document and rule."""
        if rule_name not in self.rules:
            logger.error(f"Cannot add correction: rule '{rule_name}' does not exist")
            return
            
        self.rules[rule_name].add_correction(doc_id, correct_value)
        self.save_kb()  # Save after each correction to persist learning
        
    def save_kb(self) -> None:
        """Save the knowledge base to disk."""
        try:
            with open(self.kb_path, 'wb') as f:
                pickle.dump({
                    'rules': self.rules,
                    'document_fingerprints': self.document_fingerprints
                }, f)
            logger.info(f"Knowledge base saved to {self.kb_path}")
        except Exception as e:
            logger.error(f"Failed to save knowledge base: {e}")
            
    def load_kb(self) -> None:
        """Load the knowledge base from disk if it exists."""
        if not os.path.exists(self.kb_path):
            logger.info(f"No existing knowledge base found at {self.kb_path}")
            return
            
        try:
            with open(self.kb_path, 'rb') as f:
                kb_data = pickle.load(f)
                self.rules = kb_data.get('rules', {})
                self.document_fingerprints = kb_data.get('document_fingerprints', {})
            logger.info(f"Knowledge base loaded from {self.kb_path} with {len(self.rules)} rules")
        except Exception as e:
            logger.error(f"Failed to load knowledge base: {e}")


class AdaptivePDFExtractor:
    """PDF extractor that can learn from corrections."""
    
    def __init__(self, knowledge_base_path: str = "pdf_knowledge_base.pkl"):
        """
        Initialize the extractor.
        
        Args:
            knowledge_base_path: Path to the knowledge base file
        """
        self.kb = PDFKnowledgeBase(knowledge_base_path)
        self.setup_default_rules()
        
    def setup_default_rules(self) -> None:
        """Set up default extraction rules for fund risk reports."""
        # Only add these rules if no existing KB was loaded
        if not self.kb.rules:
            # Fund report specific rules
            self.kb.add_rule(ExtractionRule(
                name="fund_name",
                pattern=r"(?:fund(?:\s+name)?|portfolio)(?:[\s\:]+)([A-Za-z0-9\s\-\&\.]+(?:Fund|Trust|ETF|Index|Portfolio))",
                description="Extract fund name",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="report_date",
                pattern=r"(?:(?:as\s+of|dated|report\s+date|date)(?:[\s\:]+)(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}))",
                description="Extract report date",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="nav",
                pattern=r"(?:NAV|Net\s+Asset\s+Value|AUM|Assets\s+Under\s+Management)(?:[\s\:]+)(?:USD|\$|€|£)?(?:\s*)(\d{1,3}(?:,\d{3})*(?:\.\d{1,9})?(?:\s*(?:million|billion|m|bn|MM|B))?)",
                description="Extract Net Asset Value",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="performance_ytd",
                pattern=r"(?:YTD|Year[\s\-]to[\s\-]Date)(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)",
                description="Extract YTD Performance",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="performance_1yr",
                pattern=r"(?:1[\s\-]?(?:Year|Yr|Y)|Annual)(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)",
                description="Extract 1-Year Performance",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="sharpe_ratio",
                pattern=r"(?:Sharpe\s+Ratio)(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,3})?)",
                description="Extract Sharpe Ratio",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="volatility",
                pattern=r"(?:Volatility|Standard\s+Deviation|Std\.\s+Dev\.)(?:[\s\:]*)(\d{1,2}(?:\.\d{1,2})?\s*\%)",
                description="Extract Volatility/Standard Deviation",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="max_drawdown",
                pattern=r"(?:Maximum\s+Drawdown|Max\s+Drawdown|Max\.\s+DD)(?:[\s\:]*)(-\d{1,2}(?:\.\d{1,2})?\s*\%)",
                description="Extract Maximum Drawdown",
                match_group=1,
            ))
            
            self.kb.add_rule(ExtractionRule(
                name="var_95",
                pattern=r"(?:Value[\s\-]at[\s\-]Risk|VaR)(?:[\s\:]*)(?:95%|@95%)?(?:[\s\:]*)(\d{1,2}(?:\.\d{1,2})?\s*\%)",
                description="Extract Value-at-Risk (95%)",
                match_group=1,
            ))
            
            self.kb.save_kb()
            
    def add_extraction_rule(self, name: str, pattern: str, description: str = None,
                          match_index: int = 0, match_group: int = 0,
                          preprocessing: List[str] = None) -> None:
        """Add a new extraction rule."""
        rule = ExtractionRule(
            name=name,
            pattern=pattern,
            description=description,
            match_index=match_index,
            match_group=match_group,
            preprocessing=preprocessing
        )
        self.kb.add_rule(rule)
        self.kb.save_kb()
        
    def extract_from_file(self, filepath: str, rules: List[str] = None) -> Dict[str, Any]:
        """
        Extract data from a PDF file using the specified rules.
        
        Args:
            filepath: Path to the PDF file
            rules: List of rule names to apply (if None, apply all rules)
            
        Returns:
            Dictionary of extracted values
        """
        # Calculate a document fingerprint
        with open(filepath, 'rb') as f:
            content = f.read(1024 * 1024)  # Read first MB for fingerprinting
            content_hash = hashlib.md5(content).hexdigest()
            
        doc_id = self.kb.get_document_id(filepath, content_hash)
        
        # Extract text from PDF
        try:
            with open(filepath, 'rb') as file:
                pdf = PdfReader(file)
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
        except Exception as e:
            logger.error(f"Error extracting text from {filepath}: {e}")
            return {'error': str(e), 'doc_id': doc_id}
            
        # Apply extraction rules
        results = {'filepath': filepath, 'doc_id': doc_id}
        
        rules_to_apply = rules or list(self.kb.rules.keys())
        for rule_name in rules_to_apply:
            rule = self.kb.get_rule(rule_name)
            if not rule:
                results[rule_name] = None
                continue
                
            value, is_correction = rule.extract(text, doc_id)
            results[rule_name] = value
            
            # Add a flag to indicate if this came from a correction
            if is_correction:
                results[f"{rule_name}_corrected"] = True
                
        return results
        
    def extract_from_directory(self, directory: str, rules: List[str] = None, 
                              output_format: str = 'csv') -> str:
        """
        Extract data from all PDFs in a directory.
        
        Args:
            directory: Directory containing PDFs
            rules: List of rule names to apply
            output_format: Format for saving results ('csv' or 'json')
            
        Returns:
            Path to the saved results file
        """
        dir_path = Path(directory)
        pdf_files = list(dir_path.glob('**/*.pdf'))
        
        results = []
        for pdf_path in tqdm(pdf_files, desc="Extracting data"):
            try:
                result = self.extract_from_file(str(pdf_path), rules)
                results.append(result)
            except Exception as e:
                logger.error(f"Error processing {pdf_path}: {e}")
                results.append({
                    'filepath': str(pdf_path),
                    'error': str(e)
                })
                
        # Save results
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if output_format.lower() == 'csv':
            output_path = f"extraction_results_{timestamp}.csv"
            pd.DataFrame(results).to_csv(output_path, index=False)
        else:
            output_path = f"extraction_results_{timestamp}.json"
            with open(output_path, 'w') as f:
                json.dump(results, f, indent=2)
                
        logger.info(f"Results saved to {output_path}")
        return output_path
        
    def add_correction(self, doc_id: str, rule_name: str, correct_value: str) -> None:
        """
        Add a correction for future extractions.
        
        Args:
            doc_id: Document ID
            rule_name: Rule name
            correct_value: The correct value
        """
        self.kb.add_correction(doc_id, rule_name, correct_value)
        
    def get_available_rules(self) -> List[Dict[str, str]]:
        """Get a list of all available extraction rules."""
        return [
            {"name": name, "description": rule.description, "pattern": rule.pattern}
            for name, rule in self.kb.rules.items()
        ]


class FundRiskReportAnalysis:
    """Advanced analysis capabilities for fund risk reports."""
    
    def __init__(self, extractor):
        """Initialize with an extractor instance."""
        self.extractor = extractor
        
    def analyze_performance_trends(self, directory: str) -> pd.DataFrame:
        """
        Analyze performance trends across multiple fund reports.
        
        Args:
            directory: Directory containing fund report PDFs
            
        Returns:
            DataFrame with performance analysis
        """
        dir_path = Path(directory)
        pdf_files = list(dir_path.glob('**/*.pdf'))
        
        performance_data = []
        for pdf_path in tqdm(pdf_files, desc="Analyzing performance"):
            try:
                result = self.extractor.extract_from_file(str(pdf_path))
                
                # Skip if missing essential data
                if not result.get('fund_name') or not result.get('report_date'):
                    continue
                    
                # Parse date
                try:
                    date_str = result.get('report_date')
                    # Try different date formats
                    for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%m-%d-%Y', '%d %b %Y', '%d %B %Y']:
                        try:
                            report_date = datetime.strptime(date_str, fmt)
                            break
                        except:
                            continue
                except:
                    # If date parsing fails, use file modification date as fallback
                    report_date = datetime.fromtimestamp(os.path.getmtime(pdf_path))
                
                # Clean up performance metrics
                perf_ytd = self._clean_percentage(result.get('performance_ytd'))
                perf_1yr = self._clean_percentage(result.get('performance_1yr'))
                volatility = self._clean_percentage(result.get('volatility'))
                max_dd = self._clean_percentage(result.get('max_drawdown'))
                
                # Create performance record
                record = {
                    'fund_name': result.get('fund_name'),
                    'report_date': report_date,
                    'nav': self._clean_nav(result.get('nav')),
                    'performance_ytd': perf_ytd,
                    'performance_1yr': perf_1yr,
                    'sharpe_ratio': self._clean_numeric(result.get('sharpe_ratio')),
                    'volatility': volatility,
                    'max_drawdown': max_dd,
                    'var_95': self._clean_percentage(result.get('var_95')),
                    'filepath': str(pdf_path)
                }
                
                performance_data.append(record)
            except Exception as e:
                logger.error(f"Error analyzing {pdf_path}: {e}")
        
        # Convert to DataFrame for analysis
        if not performance_data:
            return pd.DataFrame()
            
        df = pd.DataFrame(performance_data)
        
        # Sort by fund name and date
        if 'report_date' in df.columns:
            df = df.sort_values(['fund_name', 'report_date'])
            
        return df
    
    def _clean_percentage(self, value: str) -> float:
        """Clean percentage values, converting to float."""
        if not value:
            return None
        try:
            # Remove % sign and convert to float
            return float(value.replace('%', '').replace('+', '').strip())
        except:
            return None
            
    def _clean_numeric(self, value: str) -> float:
        """Clean numeric values, converting to float."""
        if not value:
            return None
        try:
            return float(value.strip())
        except:
            return None
            
    def _clean_nav(self, value: str) -> float:
        """Clean NAV values, handling millions/billions notation."""
        if not value:
            return None
        try:
            value = value.strip().lower()
            multiplier = 1.0
            
            # Handle million/billion notation
            if any(x in value for x in ['million', 'm', 'mm']):
                multiplier = 1_000_000
                value = value.replace('million', '').replace('m', '').replace('mm', '')
            elif any(x in value for x in ['billion', 'b', 'bn']):
                multiplier = 1_000_000_000
                value = value.replace('billion', '').replace('b', '').replace('bn', '')
                
            # Remove commas and convert to float
            value = value.replace(',', '')
            return float(value) * multiplier
        except:
            return None
    
    def generate_performance_report(self, df: pd.DataFrame, output_file: str = "fund_performance_report.csv") -> str:
        """
        Generate a performance report from analyzed data.
        
        Args:
            df: DataFrame with performance data
            output_file: Path to save the report
            
        Returns:
            Path to the saved report
        """
        if df.empty:
            logger.warning("No data available for performance report")
            return None
            
        # Calculate additional metrics if possible
        if 'performance_1yr' in df.columns and 'volatility' in df.columns:
            df['risk_adjusted_return'] = df.apply(
                lambda x: x['performance_1yr'] / x['volatility'] if x['volatility'] and x['volatility'] > 0 else None, 
                axis=1
            )
            
        # Save to file
        df.to_csv(output_file, index=False)
        logger.info(f"Performance report saved to {output_file}")
        
        return output_file
    
    def compare_funds(self, fund_names: List[str], df: pd.DataFrame) -> Dict[str, Any]:
        """
        Compare metrics for specified funds.
        
        Args:
            fund_names: List of fund names to compare
            df: DataFrame with performance data
            
        Returns:
            Dictionary with comparison results
        """
        if df.empty:
            return {'error': 'No data available for comparison'}
            
        # Filter for requested funds
        comparison_df = df[df['fund_name'].isin(fund_names)]
        
        if comparison_df.empty:
            return {'error': 'None of the specified funds found in data'}
            
        # Group by fund name and get latest report for each
        latest_reports = comparison_df.sort_values('report_date').groupby('fund_name').last().reset_index()
        
        # Create comparison dictionary
        comparison = {
            'date': datetime.now().strftime('%Y-%m-%d'),
            'funds': {}
        }
        
        metrics = ['nav', 'performance_ytd', 'performance_1yr', 'sharpe_ratio', 
                   'volatility', 'max_drawdown', 'var_95']
        
        for _, row in latest_reports.iterrows():
            fund_data = {metric: row.get(metric) for metric in metrics if metric in row}
            comparison['funds'][row['fund_name']] = fund_data
            
        return comparison


class CommandLineInterface:
    """Command line interface for the Fund Risk Report Extractor."""
    
    def __init__(self):
        self.extractor = AdaptivePDFExtractor()
        self.analyzer = FundRiskReportAnalysis(self.extractor)
        
    def run(self):
        """Run the CLI."""
        parser = argparse.ArgumentParser(description='Fund Risk Report Extractor')
        subparsers = parser.add_subparsers(dest='command', help='Command to run')
        
        # Extract command
        extract_parser = subparsers.add_parser('extract', help='Extract data from PDFs')
        extract_parser.add_argument('--input', '-i', required=True, help='Input PDF file or directory')
        extract_parser.add_argument('--rules', '-r', nargs='+', help='Rule names to apply (default: all)')
        extract_parser.add_argument('--format', '-f', choices=['csv', 'json'], default='csv', 
                                   help='Output format (default: csv)')
        
        # Rules command
        rules_parser = subparsers.add_parser('rules', help='List available rules')
        
        # Add rule command
        add_rule_parser = subparsers.add_parser('add-rule', help='Add a new extraction rule')
        add_rule_parser.add_argument('--name', required=True, help='Rule name')
        add_rule_parser.add_argument('--pattern', required=True, help='Regex pattern')
        add_rule_parser.add_argument('--description', help='Rule description')
        add_rule_parser.add_argument('--match-index', type=int, default=0, help='Match index to use')
        add_rule_parser.add_argument('--match-group', type=int, default=0, help='Match group to use')
        add_rule_parser.add_argument('--preprocessing', nargs='+', 
                                    choices=['lowercase', 'remove_whitespace'],
                                    help='Preprocessing steps')
        
        # Correct command
        correct_parser = subparsers.add_parser('correct', help='Add a correction for a document')
        correct_parser.add_argument('--doc-id', required=True, help='Document ID')
        correct_parser.add_argument('--rule', required=True, help='Rule name')
        correct_parser.add_argument('--value', required=True, help='Correct value')
        
        # Interactive command
        interactive_parser = subparsers.add_parser('interactive', help='Interactive extraction with learning')
        interactive_parser.add_argument('--input', '-i', required=True, help='Input PDF file or directory')
        
        # Fund analysis commands
        analyze_parser = subparsers.add_parser('analyze', help='Analyze fund performance across reports')
        analyze_parser.add_argument('--input', '-i', required=True, help='Directory containing fund reports')
        analyze_parser.add_argument('--output', '-o', default='fund_performance_report.csv', help='Output file for analysis')
        
        compare_parser = subparsers.add_parser('compare', help='Compare specified funds')
        compare_parser.add_argument('--input', '-i', required=True, help='Directory containing fund reports')
        compare_parser.add_argument('--funds', '-f', required=True, nargs='+', help='Fund names to compare')
        compare_parser.add_argument('--output', '-o', default='fund_comparison.json', help='Output file for comparison')
        
        args = parser.parse_args()
        
        if args.command == 'extract':
            input_path = Path(args.input)
            if input_path.is_file():
                results = self.extractor.extract_from_file(str(input_path), args.rules)
                print(json.dumps(results, indent=2))
            elif input_path.is_dir():
                output_path = self.extractor.extract_from_directory(str(input_path), args.rules, args.format)
                print(f"Results saved to {output_path}")
            else:
                print(f"Error: {input_path} is not a valid file or directory")
                
        elif args.command == 'rules':
            rules = self.extractor.get_available_rules()
            print("Available extraction rules:")
            for rule in rules:
                print(f"- {rule['name']}: {rule['description']}")
                print(f"  Pattern: {rule['pattern']}")
                print()
                
        elif args.command == 'add-rule':
            self.extractor.add_extraction_rule(
                name=args.name,
                pattern=args.pattern,
                description=args.description,
                match_index=args.match_index,
                match_group=args.match_group,
                preprocessing=args.preprocessing
            )
            print(f"Rule '{args.name}' added successfully")
            
        elif args.command == 'correct':
            self.extractor.add_correction(args.doc_id, args.rule, args.value)
            print(f"Correction added for document {args.doc_id}, rule {args.rule}")
            
        elif args.command == 'interactive':
            self._run_interactive_mode(args.input)
            
        elif args.command == 'analyze':
            print(f"Analyzing fund reports in {args.input}...")
            performance_df = self.analyzer.analyze_performance_trends(args.input)
            if performance_df.empty:
                print("No performance data could be extracted from the reports.")
            else:
                output_file = self.analyzer.generate_performance_report(performance_df, args.output)
                print(f"Analysis complete! Results saved to {output_file}")
                print(f"Found data for {performance_df['fund_name'].nunique()} funds across {len(performance_df)} reports")
                
        elif args.command == 'compare':
            print(f"Comparing funds: {', '.join(args.funds)}")
            performance_df = self.analyzer.analyze_performance_trends(args.input)
            if performance_df.empty:
                print("No performance data could be extracted from the reports.")
            else:
                comparison = self.analyzer.compare_funds(args.funds, performance_df)
                
                if 'error' in comparison:
                    print(f"Error: {comparison['error']}")
                else:
                    with open(args.output, 'w') as f:
                        json.dump(comparison, f, indent=2)
                    print(f"Comparison saved to {args.output}")
                    
                    # Print a simple comparison table
                    print("\nFund Comparison Summary:")
                    print("-" * 80)
                    metrics = ['performance_1yr', 'sharpe_ratio', 'volatility', 'max_drawdown']
                    metric_names = {
                        'performance_1yr': '1Y Return (%)', 
                        'sharpe_ratio': 'Sharpe Ratio',
                        'volatility': 'Volatility (%)', 
                        'max_drawdown': 'Max Drawdown (%)'
                    }
                    
                    # Print header
                    print(f"{'Metric':<20}", end="")
                    for fund in comparison['funds']:
                        print(f"{fund:<15}", end="")
                    print()
                    print("-" * 80)
                    
                    # Print metrics
                    for metric in metrics:
                        print(f"{metric_names.get(metric, metric):<20}", end="")
                        for fund, data in comparison['funds'].items():
                            value = data.get(metric, "N/A")
                            if value is not None:
                                if metric in ['performance_1yr', 'volatility', 'max_drawdown']:
                                    print(f"{value:>14.2f}%", end="")
                                else:
                                    print(f"{value:>14.2f}", end="")
                            else:
                                print(f"{'':<14}N/A", end="")
                        print()
            
        else:
            parser.print_help()
            
    def _run_interactive_mode(self, input_path: str) -> None:
        """
        Run in interactive mode, extracting data and learning from corrections.
        
        Args:
            input_path: PDF file or directory path
        """
        input_path = Path(input_path)
        
        if input_path.is_file():
            pdf_files = [input_path]
        elif input_path.is_dir():
            pdf_files = list(input_path.glob('**/*.pdf'))
        else:
            print(f"Error: {input_path} is not a valid file or directory")
            return
            
        rules = self.extractor.get_available_rules()
        rule_names = [rule['name'] for rule in rules]
        
        print(f"Found {len(pdf_files)} PDF files to process")
        print(f"Available rules: {', '.join(rule_names)}")
        print()
        
        for pdf_file in pdf_files:
            print(f"\nProcessing: {pdf_file}")
            results = self.extractor.extract_from_file(str(pdf_file), rule_names)
            
            print("\nExtracted values:")
            for rule_name in rule_names:
                value = results.get(rule_name)
                is_corrected = results.get(f"{rule_name}_corrected", False)
                
                if is_corrected:
                    print(f"  {rule_name}: {value} (from previous correction)")
                else:
                    print(f"  {rule_name}: {value}")
                    
                    # Ask if the value is correct
                    while True:
                        response = input(f"  Is this value correct? (y/n/s) ")
                        if response.lower() == 'y':
                            break
                        elif response.lower() == 'n':
                            correct_value = input(f"  Enter the correct value for {rule_name}: ")
                            self.extractor.add_correction(results['doc_id'], rule_name, correct_value)
                            print(f"  Correction saved for future use")
                            break
                        elif response.lower() == 's':
                            print(f"  Skipping this field")
                            break
                        else:
                            print(f"  Please enter 'y' (yes), 'n' (no), or 's' (skip)")
                            
            print("\nContinue to next file? (Enter to continue, q to quit)")
            if input().lower() == 'q':
                break
                
        print("\nInteractive session completed")


def main():
    """Main entry point for the application."""
    cli = CommandLineInterface()
    cli.run()


if __name__ == "__main__":
    main()
