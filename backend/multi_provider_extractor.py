#!/usr/bin/env python3
"""
multi_provider_extractor.py - Extract data from multiple pension providers
Supports: AJ Bell, Morningstar, and more
"""

import sys
import json
import re

def detect_provider(text):
    """Detect which provider the document is from"""
    text_lower = text.lower()
    
    if 'aj bell' in text_lower or 'ajbell' in text_lower:
        return 'AJ_BELL'
    elif 'morningstar' in text_lower:
        return 'MORNINGSTAR'
    elif 'cashflow' in text_lower or '574611' in text_lower:
        return 'CASHFLOW'
    else:
        return 'UNKNOWN'

def extract_aj_bell(text):
    """Extract data from AJ Bell performance reports"""
    
    data = {
        'provider': 'AJ Bell',
        'client_name': None,
        'accounts': [],
        'total_value': 0,
        'performance': {}
    }
    
    # Extract client name
    client_match = re.search(r'(Mr|Mrs|Ms)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)', text)
    if client_match:
        data['client_name'] = client_match.group(0).strip()
    
    # Extract account number
    acc_match = re.search(r'(SCC\d+)', text)
    account_number = acc_match.group(1) if acc_match else None
    
    # Extract ISA section
    isa_section = re.search(r'ISA - Performance summary.*?(?=ISA - Performance Analysis|ISA - Holdings|$)', text, re.DOTALL)
    if isa_section:
        isa_text = isa_section.group(0)
        
        # Get value from holdings section (Total row)
        holdings_match = re.search(r'Total\s+-\s+[\d,]+\.\d{2}\s+-\s+([\d,]+\.\d{2})', text)
        if holdings_match:
            value = float(holdings_match.group(1).replace(',', ''))
        else:
            # Fallback: look for value in summary line
            value_match = re.search(r'£([\d,]+\.\d{2})\s+[\d.]+%', isa_text)
            value = float(value_match.group(1).replace(',', '')) if value_match else 0
        
        # Get contributions (Cash in) - look in the data line
        contrib_match = re.search(r'Cash in.*?£([\d,]+\.\d{2})', text)
        contributions = float(contrib_match.group(1).replace(',', '')) if contrib_match else 0
        
        # Get performance (Time-weighted return)
        perf_match = re.search(r'([\d.]+)%\s*$', isa_text, re.MULTILINE)
        if not perf_match:
            perf_match = re.search(r'Time-weighted return.*?([\d.]+)%', text, re.DOTALL)
        performance = float(perf_match.group(1)) if perf_match else 0
        
        # Get return amount (Change in holdings)
        change_match = re.search(r'Total.*?Change.*?([\d,]+\.\d{2})', text, re.DOTALL)
        if not change_match:
            change_match = re.search(r'Change \(£\).*?([\d,]+\.\d{2})', text, re.DOTALL)
        return_amount = float(change_match.group(1).replace(',', '')) if change_match else 0
        
        account = {
            'type': 'ISA',
            'provider': 'AJ Bell',
            'account_number': account_number,
            'value': value,
            'contributions': contributions,
            'return': return_amount,
            'performance': performance
        }
        
        data['accounts'].append(account)
        data['total_value'] += value
    
    # Extract SIPP if present (similar pattern)
    sipp_section = re.search(r'SIPP - Performance summary.*?(?=SIPP - Performance Analysis|SIPP - Holdings|$)', text, re.DOTALL)
    if sipp_section:
        # Similar extraction logic for SIPP
        pass
    
    if data['accounts']:
        data['performance']['oneYearReturn'] = data['accounts'][0]['performance']
    
    return data

def extract_morningstar(text):
    """Extract data from Morningstar reports"""
    
    data = {
        'provider': 'Morningstar',
        'client_name': None,
        'accounts': [],
        'total_value': 0,
        'performance': {}
    }
    
    # Extract client name
    client_match = re.search(r'PREPARED FOR\s+((?:Ms\.|Mr\.|Mrs\.)\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)', text)
    if client_match:
        data['client_name'] = client_match.group(1).strip()
    
    # Extract ISA
    isa_section = re.search(r'Investment held:\s*ISA.*?(?=Investment held:|$)', text, re.DOTALL)
    if isa_section:
        isa_text = isa_section[0]
        
        val_match = re.search(r'Portfolio Valuation\s*£([\d,]+\.\d{2})', isa_text)
        contrib_match = re.search(r'Total In/Out\s*£([\d,]+\.\d{2})', isa_text)
        return_match = re.search(r'Total Return\s*£([\d,]+\.\d{2})', isa_text)
        perf_match = re.search(r'Portfolio % Returns\s*([\d.]+)%', isa_text)
        
        if val_match:
            account = {
                'type': 'ISA',
                'provider': 'Morningstar',
                'value': float(val_match.group(1).replace(',', '')),
                'contributions': float(contrib_match.group(1).replace(',', '')) if contrib_match else 0,
                'return': float(return_match.group(1).replace(',', '')) if return_match else 0,
                'performance': float(perf_match.group(1)) if perf_match else 0
            }
            
            data['accounts'].append(account)
            data['total_value'] += account['value']
    
    # Extract SIPP
    sipp_section = re.search(r'Investment held:\s*SIPP.*?(?=Investment held:|$)', text, re.DOTALL)
    if sipp_section:
        sipp_text = sipp_section[0]
        
        val_match = re.search(r'Portfolio Valuation\s*£([\d,]+\.\d{2})', sipp_text)
        contrib_match = re.search(r'Total In/Out\s*£([\d,]+\.\d{2})', sipp_text)
        return_match = re.search(r'Total Return\s*£([\d,]+\.\d{2})', sipp_text)
        perf_match = re.search(r'Portfolio % Returns\s*([\d.]+)%', sipp_text)
        
        if val_match:
            account = {
                'type': 'SIPP',
                'provider': 'Morningstar',
                'value': float(val_match.group(1).replace(',', '')),
                'contributions': float(contrib_match.group(1).replace(',', '')) if contrib_match else 0,
                'return': float(return_match.group(1).replace(',', '')) if return_match else 0,
                'performance': float(perf_match.group(1)) if perf_match else 0
            }
            
            data['accounts'].append(account)
            data['total_value'] += account['value']
    
    if data['accounts']:
        avg_perf = sum(acc['performance'] for acc in data['accounts']) / len(data['accounts'])
        data['performance']['oneYearReturn'] = avg_perf
    
    return data

def main():
    """Main extraction with provider detection"""
    
    # Read PDF text from stdin
    text = sys.stdin.read()
    
    # Detect provider
    provider = detect_provider(text)
    
    print(f"PROVIDER:{provider}", file=sys.stderr)
    
    # Extract based on provider
    if provider == 'AJ_BELL':
        result = extract_aj_bell(text)
    elif provider == 'MORNINGSTAR':
        result = extract_morningstar(text)
    else:
        result = {
            'provider': 'Unknown',
            'error': 'Provider not recognized',
            'accounts': [],
            'total_value': 0
        }
    
    # Output JSON
    print(json.dumps(result, indent=2))

if __name__ == '__main__':
    main()
