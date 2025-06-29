import os
import re
import json
import time
from datetime import datetime
import pandas as pd
import pdfplumber
import pytesseract
import requests
from pdf2image import convert_from_path

# --- Configuratie ---
# De werkende API sleutel uit het notebook
OPENROUTER_API_KEY = "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"
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL_NAME = "deepseek/deepseek-r1:free"

# --- Stap 1: Functies uit zebi.ipynb (PDF Analyse) ---

def extract_text_from_pdf_with_ocr(pdf_path, min_text_length=100):
    """
    Probeert eerst tekst digitaal te extraheren. Als dat te weinig tekst oplevert,
    schakelt het automatisch over naar OCR met Tesseract.
    """
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\\n"
        print("Methode: Digitale extractie (pdfplumber) geslaagd.")
    except Exception as e:
        print(f"Fout tijdens digitale extractie: {e}. Probeert nu OCR.")
        text = ""

    if len(text.strip()) >= min_text_length:
        return text

    print(f"Digitale extractie leverde < {min_text_length} tekens op. Overschakelen naar OCR...")
    try:
        images = convert_from_path(pdf_path)
        ocr_text = ""
        for i, image in enumerate(images):
            print(f"  Verwerken van pagina {i+1} met OCR...")
            ocr_text += pytesseract.image_to_string(image, lang='nld') + "\\n"
        print("Methode: OCR-extractie (Tesseract) geslaagd.")
        return ocr_text
    except Exception as e:
        print(f"Fout tijdens OCR-extractie: {e}")
        return None

def extract_percentage_sentences(text):
    """
    Zoekt naar zinnen in de tekst die percentages bevatten,
    met een focus op contexten die gerelateerd kunnen zijn aan loonstijgingen.
    (Deze functie is 1-op-1 overgenomen uit de werkende zebi.ipynb)
    """
    if not text:
        return []

    # Gecorrigeerd regex patroon. '[^.?!]' wordt gebruikt om de hele zin te vangen.
    percentage_pattern = re.compile(
        r'([^.?!]*?(?:loon|salaris|cao|verhoging|stijging|toeslag)\\s*[^.?!]*?\\d[\\d.,]*\\s?%\\s*[^.?!]*?[.?!])',
        re.IGNORECASE
    )
    sentences_with_percentage = percentage_pattern.findall(text)

    # Verwijder overtollige spaties en vervang de ECHTE newline karakters ('\\n')
    clean_sentences = [s.strip().replace('\\n', ' ') for s in sentences_with_percentage]
    
    if clean_sentences:
        print(f"[PROCESS LOG] {len(clean_sentences)} relevante zinnen gevonden voor AI-analyse.")
    else:
        print("[PROCESS LOG] WAARSCHUWING: Geen relevante zinnen gevonden na correcte regex en normalisatie.")
        
    return clean_sentences

def classify_with_deepseek(sentence, api_key, api_url, model_name, max_retries=3, delay=5):
    """
    Roept de API aan via OpenRouter om een zin te classificeren en om uitleg te vragen.
    """
    if not api_key:
        print("Fout: OpenRouter API-sleutel is leeg of niet ingesteld.")
        return None

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://localhost/sakkal-cao-analyzer",
        "X-Title": "Team Sakkal CAO Analyzer"
    }
    
    system_prompt = """Je bent een expert in het analyseren van Nederlandse CAO-teksten. Je taak is om te bepalen of een zin concrete, definitieve loonstijgingen beschrijft en **alle** details daarvan te extraheren.

**Instructies:**
- Een zin kan **meerdere** loonstijgingen bevatten (bijv. in een opsomming). Je moet ze **allemaal** vinden.
- Classificeer alleen **daadwerkelijke, gegarandeerde afspraken** over salarisverhogingen als 'Loonstijging'.
- **Negeer** zinnen die beginnen met 'Rekenvoorbeeld', 'Voorbeeld:', 'Stel dat', of 'Berekening'.
- **Negeer** structurele salarisverschillen tussen functies (bv. 'Level 2 is 10% hoger dan Level 1').
- **Cruciaal: Negeer voorwaardelijke verhogingen** die afhankelijk zijn van een conditie (bv. 'tenzij', 'indien').
- **Cruciaal: Negeer vergelijkende berekeningen.** Als een stijging wordt beschreven 'ten opzichte van' een datum in het verleden, is dit een vergelijking, geen nieuwe afspraak.

**Jouw taak:**
Analyseer de zin en geef een JSON-object terug met DRIE sleutels:
1. 'classificatie': 'Loonstijging' of 'Geen Loonstijging'.
2. 'verhogingen': Een **lijst** van JSON-objecten. Elk object bevat 'datum' en 'percentage'. Als er geen gegarandeerde loonstijgingen zijn, moet deze lijst leeg zijn: `[]`.
3. 'uitleg': een korte, duidelijke toelichting op je keuze.

**BELANGRIJK:** Je antwoord MOET **uitsluitend** een enkel, geldig JSON-object zijn. Geen extra tekst."""

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "De salarissen stijgen met 2% op 01-01-2025 en met nog eens 3% op 01-07-2025."},
        {"role": "assistant", "content": '''{"classificatie": "Loonstijging", "verhogingen": [{"datum": "01/01/2025", "percentage": 2.0}, {"datum": "01/07/2025", "percentage": 3.0}], "uitleg": "De zin bevat twee concrete loonstijgingen."}'''},
        {"role": "user", "content": "Met ingang van 1 januari 2025 worden deze verlofdagen omgezet in een structurele salarisverhoging van 0,84% tenzij werkgever deze verlofdagen al eerder heeft omgezet."},
        {"role": "assistant", "content": '''{"classificatie": "Geen Loonstijging", "verhogingen": [], "uitleg": "De verhoging is voorwaardelijk vanwege de 'tenzij'-clausule en dus niet gegarandeerd."}'''},
        {"role": "user", "content": "het eindloon stijgt ten opzichte van het eindloon op 31 december 2024 met 6,9%."},
        {"role": "assistant", "content": '''{"classificatie": "Geen Loonstijging", "verhogingen": [], "uitleg": "Dit is een vergelijkende berekening ten opzichte van een referentiedatum, geen nieuwe loonafspraak met een ingangsdatum."}'''},
        {"role": "user", "content": sentence}
    ]

    data = { "model": model_name, "messages": messages, "response_format": { "type": "json_object" }, "max_tokens": 400, "temperature": 0.0 }
    
    for attempt in range(max_retries):
        try:
            print(f"[AI LOG] Attempt {attempt + 1}: Sending request to model '{model_name}'...")
            response = requests.post(api_url, json=data, headers=headers, timeout=45)
            print(f"[AI LOG] Received response with status code: {response.status_code}")
            
            # For debugging, always print the raw text
            raw_response_text = response.text
            print(f"[AI LOG] Raw AI Response:\\n---\\n{raw_response_text}\\n---")
            
            response.raise_for_status()
            
            try:
                content = response.text
                full_response_json = json.loads(content)
                message_content = full_response_json['choices'][0]['message']['content']
                
                start_index = message_content.rfind('{')
                end_index = message_content.rfind('}')
                
                if start_index != -1 and end_index != -1 and end_index > start_index:
                    json_str = message_content[start_index : end_index + 1]
                    try:
                        return json.loads(json_str)
                    except json.JSONDecodeError as e:
                        print(f"Fout bij parsen van laatst gevonden JSON-object: {e}")
                        print(f"Geïsoleerde tekst: {json_str}")
                        return None
                else:
                    print("Geen JSON-object gevonden in de AI-respons.")
                    print(f"Ontvangen tekst: {message_content}")
                    return None
            except (json.JSONDecodeError, KeyError, IndexError, TypeError) as e:
                print(f"Fout bij parsen van de volledige API-respons: {e}")
                print(f"Ontvangen tekst: {response.text}")
                return None
        except requests.exceptions.RequestException as e:
            print(f"Netwerkfout: {e}")
            if attempt < max_retries - 1:
                time.sleep(delay)
            else:
                return None
    
    print(f"Fout na {max_retries} pogingen.")
    return None

def analyze_pdfs(pdf_paths):
    """
    Orchestreert het volledige analyseproces voor een lijst van PDF-bestanden.
    """
    final_json_output = {}
    for pdf_path in pdf_paths:
        pdf_filename = os.path.basename(pdf_path)
        print("\\n" + "="*80)
        print(f"--- Begin analyse van: {pdf_filename} ---\\n")
        
        extracted_text = extract_text_from_pdf_with_ocr(pdf_path)
        
        if not extracted_text:
            print(f"Kon geen tekst extraheren uit {pdf_filename}. Bestand wordt overgeslagen.")
            final_json_output[pdf_filename] = {"error": "Tekstextractie mislukt", "verhogingen": []}
            continue

        sentences = extract_percentage_sentences(extracted_text)
        if not sentences:
            print(f"Geen relevante zinnen met percentages gevonden in {pdf_filename}.")
            final_json_output[pdf_filename] = {"error": "Geen relevante zinnen gevonden", "verhogingen": []}
            continue

        print(f"{len(sentences)} zinnen gevonden voor analyse.")
        
        all_found_increases = []
        print("\\n[PROCESS LOG] Beginnen met classificatie via DeepSeek API:")
        for i, sentence in enumerate(sentences):
            short_sentence = (sentence[:120] + '...') if len(sentence) > 120 else sentence
            print(f"\\n[PROCESS LOG] Verwerken zin {i+1}/{len(sentences)}: {short_sentence}")
            
            result_json = classify_with_deepseek(sentence, OPENROUTER_API_KEY, OPENROUTER_API_URL, MODEL_NAME)
            print(f"[PROCESS LOG] Parsed JSON from AI: {result_json}")
            
            if result_json:
                classificatie = result_json.get('classificatie', 'Onbekend')
                verhogingen = result_json.get('verhogingen', [])
                
                if classificatie == 'Loonstijging' and verhogingen:
                    print(f"  Resultaat: ✅ {len(verhogingen)} loonstijging(en) gevonden.")
                    all_found_increases.extend(verhogingen)
                else:
                    print(f"  Resultaat: ❌ Geen Loonstijging")
            else:
                print("  Resultaat: ❓ Fout bij het classificeren van de zin.")
            
            time.sleep(1)

        # Stap 4: Resultaten voor de huidige PDF groeperen en aggregeren
        grouped_increases = {}
        for increase in all_found_increases:
            datum = increase.get('datum')
            percentage_val = increase.get('percentage')

            # Controleer of de data valide is en niet 'N.v.t.'
            if datum and percentage_val and str(datum).lower() != 'n.v.t.' and str(percentage_val).lower() != 'n.v.t.':
                parsed_percentage = None
                # Probeer de waarde om te zetten naar een getal (float)
                if isinstance(percentage_val, (int, float)):
                    parsed_percentage = float(percentage_val)
                elif isinstance(percentage_val, str):
                    try:
                        # Maak de string schoon: verwijder '%' en vervang ',' door '.'
                        cleaned_val = percentage_val.strip().replace('%', '').replace(',', '.').strip()
                        if cleaned_val:  # Zorg dat er een waarde overblijft
                            parsed_percentage = float(cleaned_val)
                    except ValueError:
                        print(f"Waarschuwing: Kon percentage-string niet omzetten naar getal: '{percentage_val}'")
                        pass  # Sla over als conversie mislukt

                # Als de conversie succesvol was, voeg toe aan de lijst
                if parsed_percentage is not None:
                    if datum not in grouped_increases:
                        grouped_increases[datum] = []
                    grouped_increases[datum].append(parsed_percentage)

        aggregated_increases = []
        for datum, percentages in grouped_increases.items():
            # Sorteer percentages om de output voorspelbaar te maken
            percentages.sort()
            percentage_str = "/".join([f"{p:.2f}".replace('.', ',') + '%' for p in percentages])
            aggregated_increases.append({'datum': datum, 'percentage': percentage_str})
        
        def sort_key(item):
            try:
                return datetime.strptime(item.get('datum', ''), '%d/%m/%Y')
            except (ValueError, TypeError):
                return datetime.max
        aggregated_increases.sort(key=sort_key)
        
        final_json_output[pdf_filename] = {"verhogingen": aggregated_increases}
        print(f"\\nAnalyse voor {pdf_filename} voltooid.")
        
    return final_json_output

# --- Stap 2: Functie uit zebiexcel.ipynb (Excel Generatie) ---

def create_excel_summary(json_data, output_path):
    """
    Verwerkt de JSON-data en genereert een Excel-bestand.
    """
    processed_data = []
    for filename, data in json_data.items():
        row_data = {}
        row_data['Cao-code'] = filename[:4]
        
        increases = data.get('verhogingen', [])
        
        for i, increase in enumerate(increases):
            date_col_name = f"{i+1}e datum nieuwe cao"
            perc_col_name = f"{i+1}e percentage nieuwe cao"
            
            row_data[date_col_name] = increase.get('datum')
            row_data[perc_col_name] = increase.get('percentage')
            
        processed_data.append(row_data)

    column_order = ['Cao-code']
    for i in range(1, 8): # Maximaal 7 verhogingen
        column_order.append(f"{i}e datum nieuwe cao")
        column_order.append(f"{i}e percentage nieuwe cao")

    df = pd.DataFrame(processed_data)

    for col in column_order:
        if col not in df.columns:
            df[col] = None

    df = df[column_order]

    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    df.to_excel(writer, index=False)
    writer.close()
    
    print(f"Excel-bestand succesvol aangemaakt op: {output_path}")

# --- Hoofd execution block ---

if __name__ == '__main__':
    print("--- Script wordt direct uitgevoerd voor testdoeleinden ---")
    
    # Gebruik een van de PDF's uit de notebook voor de test
    test_pdf_paths = [
        "/Users/mohamedelharchaoui/Downloads/Assignment 3/0083_01-01-2025 tot 01-07-2026_akkoord.pdf"
        # Voeg hier eventueel meer paden toe om te testen
    ]
    
    # Controleer of de API key is ingesteld
    if not OPENROUTER_API_KEY:
        print("FOUT: OPENROUTER_API_KEY is niet ingesteld. Controleer de configuratie bovenaan het script.")
    else:
        # Voer de analyse uit
        analysis_result = analyze_pdfs(test_pdf_paths)
        
        # Print de volledige, onbewerkte JSON-output
        print("\\n" + "="*80)
        print("--- Volledige JSON-Resultaat van de Analyse ---")
        print(json.dumps(analysis_result, indent=2, ensure_ascii=False))
        print("="*80)

        # Optioneel: Genereer ook een Excel-bestand voor de test
        if analysis_result and not any(v.get("error") for v in analysis_result.values()):
             # Definieer een pad voor het test-excelbestand
            test_excel_path = "test_output.xlsx"
            create_excel_summary(analysis_result, test_excel_path) 