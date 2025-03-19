#!/usr/bin/env python3
"""
Fund Report Extractor - A tool for extracting data from fund risk reports with learning capabilities.

Usage:
    1. Install dependencies: pip install PyPDF2 pandas tqdm colorama
    2. Run: python fund_report_extractor.py
    3. Use the GUI to select and analyze your fund report PDFs

This script contains a complete, self-contained implementation that can be run directly.
"""

import os
import re
import hashlib
import pickle
import json
import argparse
import threading
import queue
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional, Union

# Import required libraries
try:
    import pandas as pd
    from PyPDF2 import PdfReader
    from tqdm import tqdm
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    
    # Try to import colorama for colored CLI output
    try:
        from colorama import init, Fore, Style
        init()  # Initialize colorama
        COLOR_SUPPORT = True
    except ImportError:
        # Define dummy color classes
        class DummyFore:
            RED = ""
            GREEN = ""
            YELLOW = ""
            CYAN = ""
            MAGENTA = ""
        class DummyStyle:
            RESET_ALL = ""
        Fore = DummyFore()
        Style = DummyStyle()
        COLOR_SUPPORT = False
        
except ImportError as e:
    missing_lib = str(e).split("'")[1]
    print(f"Error: Missing required library: {missing_lib}")
    print("Please install required libraries with:")
    print("pip install PyPDF2 pandas tqdm colorama")
    exit(1)

# Configure logging
import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger('fund_report_extractor')

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
    
    def __init__(self, kb_path: str = "fund_report_knowledge_base.pkl"):
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
                content = f.read(1024 * 1024)  # Read first MB for fingerprinting
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


class FundReportExtractor:
    """Interactive extractor for fund reports with learning capabilities and customizable metrics."""
    
    def __init__(self, knowledge_base_path: str = "fund_report_knowledge_base.pkl"):
        """
        Initialize the extractor.
        
        Args:
            knowledge_base_path: Path to the knowledge base file
        """
        self.kb = PDFKnowledgeBase(knowledge_base_path)
        self.setup_default_rules()
        self.current_results = []  # Store current extraction results
        
    def setup_default_rules(self) -> None:
        """Set up default extraction rules for fund risk reports.
        These are just starting points - users can add their own rules.
        """
        # Only add these rules if no existing KB was loaded
        if not self.kb.rules:
            # Fund report common rules - these are just suggestions, not exhaustive
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
            
            # These are just examples - the system will prompt users to add more metrics specific to their reports
            self.kb.save_kb()
            
    def prompt_for_custom_metrics(self, sample_text: str = None) -> List[str]:
        """
        Prompt the user to add custom metrics they want to extract.
        Optionally shows them a sample of text from their document.
        
        Args:
            sample_text: Sample text from a document to help the user identify metrics
            
        Returns:
            List of added metric names
        """
        print("\n=== Custom Metrics Configuration ===")
        print("Let's configure additional metrics to extract from your fund reports.")
        
        if sample_text:
            print("\nHere's a sample from your document to help identify metrics:")
            print("-" * 80)
            # Show first 500 chars of sample text
            print(sample_text[:500] + "..." if len(sample_text) > 500 else sample_text)
            print("-" * 80)
        
        print("\nCommon fund metrics include:")
        print("- NAV (Net Asset Value)")
        print("- Performance metrics (YTD, 1Y, 3Y, 5Y returns)")
        print("- Risk metrics (Sharpe ratio, volatility, max drawdown)")
        print("- Portfolio statistics (number of holdings, top positions)")
        print("- Expense ratios and fees")
        print("- Benchmark comparisons")
        print("- Portfolio manager information")
        
        added_metrics = []
        
        print("\nLet's add the metrics you want to extract.")
        print("For each metric, you'll provide a name, description, and pattern.")
        print("Enter 'done' when finished adding metrics.")
        
        while True:
            metric_name = input("\nMetric name (or 'done' to finish): ").strip()
            if metric_name.lower() == 'done':
                break
                
            description = input(f"Description for {metric_name}: ").strip()
            
            # Suggest a pattern based on the metric name
            suggested_pattern = self._suggest_pattern(metric_name)
            
            print(f"\nSuggested pattern: {suggested_pattern}")
            pattern_input = input("Use this pattern? (y/n): ").strip().lower()
            
            if pattern_input == 'y':
                pattern = suggested_pattern
            else:
                print("\nEnter a regular expression pattern to extract this metric.")
                print("Example: 'NAV(?:[\s\:]+)(\d+(?:\.\d+)?(?:\s*(?:million|billion|m|bn)?))' to extract NAV values.")
                pattern = input("Pattern: ").strip()
                
            # Add the rule
            match_group = int(input("Capture group to use (usually 1): ").strip() or "1")
            
            self.kb.add_rule(ExtractionRule(
                name=metric_name,
                pattern=pattern,
                description=description,
                match_group=match_group
            ))
            
            added_metrics.append(metric_name)
            print(f"\n{metric_name} added successfully!")
            
        if added_metrics:
            self.kb.save_kb()
            print(f"\nAdded {len(added_metrics)} custom metrics: {', '.join(added_metrics)}")
        
        return added_metrics
        
    def _suggest_pattern(self, metric_name: str) -> str:
        """Suggest a regex pattern based on the metric name."""
        # Convert to lowercase and remove spaces for comparison
        metric_lower = metric_name.lower().replace(" ", "")
        
        # Performance metrics
        if any(term in metric_lower for term in ['return', 'performance', 'yield']):
            if 'ytd' in metric_lower or 'year-to-date' in metric_lower:
                return r"(?:YTD|Year[\s\-]to[\s\-]Date)(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)"
            elif '1y' in metric_lower or '1-year' in metric_lower or 'one-year' in metric_lower:
                return r"(?:1[\s\-]?(?:Year|Yr|Y)|Annual)(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)"
            elif '3y' in metric_lower or '3-year' in metric_lower:
                return r"(?:3[\s\-]?(?:Year|Yr|Y))(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)"
            elif '5y' in metric_lower or '5-year' in metric_lower:
                return r"(?:5[\s\-]?(?:Year|Yr|Y))(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)"
            elif '10y' in metric_lower or '10-year' in metric_lower:
                return r"(?:10[\s\-]?(?:Year|Yr|Y))(?:[\s\:]*)(?:Return|Performance|Change)?(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)"
            else:
                return r"(?:" + metric_name + r")(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,2})?\s*\%)"
        
        # NAV / AUM metrics
        elif 'nav' in metric_lower or 'aum' in metric_lower or 'asset' in metric_lower:
            return r"(?:" + metric_name + r")(?:[\s\:]+)(?:USD|\$|€|£)?(?:\s*)(\d{1,3}(?:,\d{3})*(?:\.\d{1,9})?(?:\s*(?:million|billion|m|bn|MM|B))?)"
        
        # Risk metrics
        elif 'sharp' in metric_lower or 'ratio' in metric_lower:
            return r"(?:Sharpe\s+Ratio)(?:[\s\:]*)([+-]?\d{1,2}(?:\.\d{1,3})?)"
        elif 'volatility' in metric_lower or 'deviation' in metric_lower:
            return r"(?:" + metric_name + r")(?:[\s\:]*)(\d{1,2}(?:\.\d{1,2})?\s*\%)"
        elif 'drawdown' in metric_lower:
            return r"(?:Maximum\s+Drawdown|Max\s+Drawdown|Max\.\s+DD)(?:[\s\:]*)(-\d{1,2}(?:\.\d{1,2})?\s*\%)"
        elif 'var' in metric_lower or 'value-at-risk' in metric_lower:
            return r"(?:Value[\s\-]at[\s\-]Risk|VaR)(?:[\s\:]*)(?:95%|@95%)?(?:[\s\:]*)(\d{1,2}(?:\.\d{1,2})?\s*\%)"
        
        # Expense metrics
        elif 'expense' in metric_lower or 'fee' in metric_lower or 'ratio' in metric_lower:
            return r"(?:" + metric_name + r")(?:[\s\:]*)(\d{1,2}(?:\.\d{1,2})?\s*\%)"
        
        # Manager metrics
        elif 'manager' in metric_lower:
            return r"(?:Portfolio\s+Manager|Fund\s+Manager|Managed\s+by)(?:[\s\:]+)([A-Za-z\.\s]+)"
        
        # Holdings metrics
        elif 'holding' in metric_lower or 'position' in metric_lower:
            return r"(?:Number\s+of\s+Holdings|Positions)(?:[\s\:]+)(\d+)"
        
        # Benchmark metrics
        elif 'benchmark' in metric_lower or 'index' in metric_lower:
            return r"(?:Benchmark|Index|Compared\s+to)(?:[\s\:]+)([A-Za-z0-9\s\&\.\-]+)"
        
        # Default pattern - just look for the metric name followed by some text
        return r"(?:" + metric_name + r")(?:[\s\:]+)([\w\s\.\,\%\$\€\£\-\+]+)"
        
    def extract_from_file(self, filepath: str, rules: List[str] = None) -> Dict[str, Any]:
        """
        Extract data from a PDF file using the specified rules.
        
        Args:
            filepath: Path to the PDF file
            rules: List of rule names to apply (if None, apply all rules)
            
        Returns:
            Dictionary of extracted values
        """
        try:
            # Calculate a document fingerprint
            with open(filepath, 'rb') as f:
                content = f.read(1024 * 1024)  # Read first MB for fingerprinting
                content_hash = hashlib.md5(content).hexdigest()
                
            doc_id = self.kb.get_document_id(filepath, content_hash)
            
            # Extract text from PDF
            with open(filepath, 'rb') as file:
                pdf = PdfReader(file)
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
                    
            # If this is the first extraction and there are no custom metrics,
            # ask the user if they want to add custom metrics
            if not rules and len(self.kb.rules) <= 2:  # Only basic rules exist
                print(f"\n{Fore.CYAN if COLOR_SUPPORT else ''}It looks like you haven't configured custom metrics yet.{Style.RESET_ALL if COLOR_SUPPORT else ''}")
                if input("Would you like to add custom metrics to extract? (y/n): ").lower() == 'y':
                    self.prompt_for_custom_metrics(text)
            
            # Apply extraction rules
            results = {'filepath': filepath, 'doc_id': doc_id, 'filename': os.path.basename(filepath)}
            
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
        except Exception as e:
            logger.error(f"Error extracting from {filepath}: {e}")
            return {'filepath': filepath, 'error': str(e)}
        
    def extract_from_files(self, filepaths: List[str], rules: List[str] = None) -> List[Dict[str, Any]]:
        """
        Extract data from multiple PDF files.
        
        Args:
            filepaths: List of PDF file paths
            rules: List of rule names to apply
            
        Returns:
            List of dictionaries with extracted values
        """
        results = []
        for filepath in filepaths:
            try:
                result = self.extract_from_file(filepath, rules)
                results.append(result)
            except Exception as e:
                logger.error(f"Error processing {filepath}: {e}")
                results.append({
                    'filepath': filepath,
                    'filename': os.path.basename(filepath),
                    'error': str(e)
                })
                
        self.current_results = results
        return results
        
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
            {"name": rule_name, "description": rule.description, "pattern": rule.pattern}
            for rule_name, rule in self.kb.rules.items()
        ]
        
    def add_rule(self, name: str, pattern: str, description: str = None,
                match_index: int = 0, match_group: int = 0,
                preprocessing: List[str] = None) -> bool:
        """
        Add a new extraction rule.
        
        Args:
            name: Rule name
            pattern: Regex pattern
            description: Rule description
            match_index: Match index to use
            match_group: Match group to use
            preprocessing: Preprocessing steps
            
        Returns:
            True if rule was added successfully
        """
        try:
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
            return True
        except Exception as e:
            logger.error(f"Error adding rule: {e}")
            return False
            
    def save_results_to_csv(self, results: List[Dict[str, Any]], filepath: str) -> bool:
        """
        Save extraction results to a CSV file.
        
        Args:
            results: List of extraction results
            filepath: Output file path
            
        Returns:
            True if saved successfully
        """
        try:
            df = pd.DataFrame(results)
            df.to_csv(filepath, index=False)
            logger.info(f"Results saved to {filepath}")
            return True
        except Exception as e:
            logger.error(f"Error saving results: {e}")
            return False
            
    def analyze_performance(self, results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Analyze performance data across multiple reports.
        
        Args:
            results: List of extraction results
            
        Returns:
            Dictionary with analysis results
        """
        analysis = {
            'total_reports': len(results),
            'funds': {}
        }
        
        # Group by fund name
        for result in results:
            fund_name = result.get('fund_name')
            if not fund_name:
                continue
                
            if fund_name not in analysis['funds']:
                analysis['funds'][fund_name] = []
                
            analysis['funds'][fund_name].append(result)
            
        # Calculate statistics
        for fund_name, fund_results in analysis['funds'].items():
            # Sort by date if possible
            try:
                fund_results.sort(key=lambda x: x.get('report_date', ''))
            except:
                pass
                
            # Get latest report
            latest_report = fund_results[-1] if fund_results else None
            
            # Extract metrics from latest report
            if latest_report:
                metrics = {}
                for field, value in latest_report.items():
                    if field not in ['filepath', 'doc_id', 'filename', 'error', 'fund_name'] and not field.endswith('_corrected'):
                        metrics[field] = value
                
                analysis['funds'][fund_name] = {
                    'reports_count': len(fund_results),
                    'latest_report': latest_report.get('report_date'),
                    **metrics
                }
            
        return analysis


class InteractiveFundExtractorApp:
    """Interactive GUI application for fund report extraction with custom metrics support."""
    
    def __init__(self, root):
        """Initialize the application."""
        self.root = root
        self.root.title("Interactive Fund Report Extractor")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        
        self.extractor = FundReportExtractor()
        self.selected_files = []
        self.extraction_results = []
        
        self.create_widgets()
        self.setup_chat_interface()
        
        # Message queue for the chat interface
        self.msg_queue = queue.Queue()
        self.processing = False
        
        # Start the message processing thread
        threading.Thread(target=self.process_messages, daemon=True).start()
        
        # Welcome message
        self.add_system_message("Welcome to the Interactive Fund Report Extractor!")
        self.add_system_message("To get started, click 'Select Files' to choose PDF files to analyze.")
        
        # Create menus
        self.create_menus()
        
    def create_menus(self):
        """Create application menus."""
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Select Files", command=self.select_files)
        file_menu.add_command(label="Process Files", command=self.process_files)
        file_menu.add_separator()
        file_menu.add_command(label="Export Results", command=self.export_results)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Metrics menu
        metrics_menu = tk.Menu(menubar, tearoff=0)
        metrics_menu.add_command(label="Configure Metrics", command=self.show_metrics_config)
        metrics_menu.add_command(label="View Available Metrics", command=self.show_available_metrics)
        metrics_menu.add_command(label="Add New Metric", command=self.add_new_metric)
        menubar.add_cascade(label="Metrics", menu=metrics_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Help", command=self.show_help)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menubar)
        
    def create_widgets(self):
        """Create the main application widgets."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top frame for file selection
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File selection button
        select_btn = ttk.Button(top_frame, text="Select Files", command=self.select_files)
        select_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Selected files counter
        self.files_label = ttk.Label(top_frame, text="No files selected")
        self.files_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Process button
        self.process_btn = ttk.Button(top_frame, text="Extract Data", command=self.process_files, state=tk.DISABLED)
        self.process_btn.pack(side=tk.RIGHT)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create chat tab
        self.chat_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.chat_frame, text="Interactive Chat")
        
        # Create results tab
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="Results Table")
        
        # Create table in results tab
        self.create_results_table()
        
        # Create export frame
        export_frame = ttk.Frame(main_frame)
        export_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Export button
        self.export_btn = ttk.Button(export_frame, text="Export Results", command=self.export_results, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT)
