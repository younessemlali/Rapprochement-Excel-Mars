import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta
import re
from typing import Dict, List, Tuple, Optional
import logging
from rapidfuzz import fuzz
import unidecode

# Configuration de la page
st.set_page_config(
    page_title="Excel Matcher Hybride",
    page_icon="🔗",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def normalize_text(s):
    """Normalise le texte comme Copilot"""
    if pd.isna(s) or s is None:
        return ""
    return unidecode.unidecode(str(s)).replace(" ", "").lower()

class HybridExcelMatcher:
    """Classe hybride pour le rapprochement Excel PDF et Beeline"""
    
    def __init__(self):
        self.excel_pdf_data = []
        self.excel_beeline_data = []
        self.matched_data = []
        self.unmatched_pdf = []
        self.unmatched_beeline = []
        self.matching_stats = {}
        self.matching_method = None
    
    def safe_float(self, value) -> Optional[float]:
        """Convertit une valeur en float de manière sécurisée"""
        if value is None or pd.isna(value):
            return None
        
        try:
            if isinstance(value, str):
                value = re.sub(r'[^\d\.,\-]', '', str(value))
                if not value:
                    return None
                
                if ',' in value and '.' not in value:
                    value = value.replace(',', '.')
                elif ',' in value and '.' in value:
                    value = value.replace(',', '')
            
            return float(value)
        except (ValueError, TypeError):
            return None
    
    def normalize_commande_number(self, commande_str) -> Optional[str]:
        """Normalise les numéros de commande"""
        if not commande_str:
            return None
        
        commande_clean = str(commande_str).strip()
        commande_clean = re.sub(r'[^\d]', '', commande_clean)
        
        return commande_clean if commande_clean else None
    
    def extract_commande_from_filename(self, filename: str) -> Optional[str]:
        """Extrait le numéro de commande depuis le nom du fichier PDF"""
        try:
            parts = filename.split('_')
            
            for part in parts:
                if len(part) == 10 and part.startswith('56') and part.isdigit():
                    return part
            
            commande_patterns = [
                r'(56\d{8})',
                r'(\d{10})',
            ]
            
            for pattern in commande_patterns:
                matches = re.findall(pattern, filename)
                if matches:
                    return matches[0]
            
            return None
            
        except Exception as e:
            logger.warning(f"Impossible d'extraire commande de {filename}: {e}")
            return None
    
    def load_excel_pdf_files(self, uploaded_files) -> List[Dict]:
        """Charge les fichiers Excel issus de l'App 1"""
        all_data = []
        
        for uploaded_file in uploaded_files:
            st.write(f"📊 Traitement Excel PDF: {uploaded_file.name}")
            
            try:
                excel_data = pd.read_excel(uploaded_file, sheet_name=None)
                
                priority_sheets = ['Donnees_Analyse', 'Résumé_Factures', 'Detail_Lignes', 'Analyse_Rubriques']
                sheets_to_process = []
                
                for sheet_name in priority_sheets:
                    if sheet_name in excel_data:
                        sheets_to_process.append(sheet_name)
                
                for sheet_name in excel_data.keys():
                    if sheet_name not in sheets_to_process:
                        sheets_to_process.append(sheet_name)
                
                total_rows_added = 0
                
                for sheet_name in sheets_to_process:
                    df_to_use = excel_data[sheet_name]
                    
                    if df_to_use is None or len(df_to_use) == 0:
                        continue
                    
                    df_cleaned = df_to_use.dropna(how='all')
                    df_cleaned.columns = df_cleaned.columns.str.strip()
                    
                    has_useful_columns = any(col in df_cleaned.columns for col in 
                                           ['Numero_Facture', 'Numero_Commande', 'N° Facture', 'N° Commande'])
                    
                    if not has_useful_columns:
                        continue
                    
                    rows_added_sheet = 0
                    for _, row in df_cleaned.iterrows():
                        # Extraire le vrai numéro de commande du nom de fichier
                        true_commande = self.extract_commande_from_filename(uploaded_file.name)
                        
                        data_row = {
                            'source_file': uploaded_file.name,
                            'source_sheet': sheet_name,
                            'numero_facture': row.get('Numero_Facture') or row.get('N° Facture'),
                            'numero_commande': true_commande or self.normalize_commande_number(row.get('Numero_Commande') or row.get('N° Commande')),
                            'date_facture': row.get('Date_Facture') or row.get('Date'),
                            'semaine_finissant_le': row.get('Semaine_Finissant_Le') or row.get('Date_Periode'),
                            'destinataire': row.get('Destinataire'),
                            'batch_id': row.get('Batch_ID'),
                            'assignment_id': row.get('Assignment_ID'),
                            'total_net': self.safe_float(row.get('Total_Net_EUR') or row.get('Total_Net') or row.get('Montant_Net')),
                            'total_tva': self.safe_float(row.get('Total_TVA_EUR') or row.get('Total_TVA') or row.get('Montant_TVA')),
                            'total_brut': self.safe_float(row.get('Total_Brut_EUR') or row.get('Total_Brut') or row.get('Montant_Brut')),
                            'code_rubrique': row.get('Code_Rubrique'),
                            'type_prestation': row.get('Type_Prestation'),
                            'type_donnees': 'PDF_EXTRACT'
                        }
                        
                        if data_row['numero_commande']:
                            all_data.append(data_row)
                            rows_added_sheet += 1
                    
                    total_rows_added += rows_added_sheet
                
                st.write(f"   ✅ {total_rows_added} lignes ajoutées du fichier {uploaded_file.name}")
                
            except Exception as e:
                st.error(f"❌ Erreur lors du traitement de {uploaded_file.name}: {e}")
        
        self.excel_pdf_data = all_data
        return all_data
    
    def load_excel_beeline_files(self, uploaded_files) -> List[Dict]:
        """Charge les fichiers Excel Beeline"""
        all_data = []
        
        for uploaded_file in uploaded_files:
            st.write(f"📋 Traitement Excel Beeline: {uploaded_file.name}")
            
            try:
                df = pd.read_excel(uploaded_file)
                df_cleaned = df.dropna(how='all')
                df_cleaned.columns = df_cleaned.columns.str.strip()
                
                rows_added = 0
                for _, row in df_cleaned.iterrows():
                    data_row = {
                        'source_file': uploaded_file.name,
                        'collaborateur': row.get('Collaborateur'),
                        'numero_commande': self.normalize_commande_number(row.get('N° commande') or row.get('Numero_Commande')),
                        'semaine_finissant_le': row.get('Semaine finissant le') or row.get('Semaine_Finissant_Le'),
                        'code_rubrique': row.get('Code rubrique') or row.get('Code_Rubrique'),
                        'taux_facturation': self.safe_float(row.get('Taux de facturation')),
                        'unites': self.safe_float(row.get('Unités')),
                        'montant_brut': self.safe_float(row.get('Montant brut')),
                        'montant_net_fournisseur': self.safe_float(row.get('Montant net à payer au fournisseur')),
                        'supplier': row.get('Supplier'),
                        'projet': row.get('Projet'),
                        'centre_cout': row.get('Centre de coût'),
                        'invoice_number': row.get('Invoice Number'),
                        'billing_period': row.get('Billing Period'),
                        'type_donnees': 'BEELINE'
                    }
                    
                    if data_row['numero_commande'] and data_row['collaborateur']:
                        all_data.append(data_row)
                        rows_added += 1
                
                st.write(f"   ✅ {rows_added} lignes ajoutées du fichier {uploaded_file.name}")
                
            except Exception as e:
                st.error(f"❌ Erreur lors du traitement de {uploaded_file.name}: {e}")
        
        self.excel_beeline_data = all_data
        return all_data
    
    def fuzzy_match_beeline_row(self, pdf_row: Dict, beeline_candidates: List[Dict], threshold: int = 85) -> Tuple[Optional[Dict], int, List[str]]:
        """Fait un matching fuzzy inspiré de Copilot entre une ligne PDF et des candidats Beeline"""
        
        best_match = None
        best_score = 0
        best_matched_fields = []
        
        # Données de la ligne PDF pour comparaison
        pdf_commande = normalize_text(str(pdf_row.get('numero_commande', '')))
        pdf_total_net = pdf_row.get('total_net', 0) or 0
        pdf_rubrique = normalize_text(str(pdf_row.get('code_rubrique', '')))
        pdf_batch = normalize_text(str(pdf_row.get('batch_id', '')))
        
        for beeline_row in beeline_candidates:
            score = 0
            matched_fields = []
            
            # Correspondance numéro de commande (obligatoire)
            beeline_commande = normalize_text(str(beeline_row.get('numero_commande', '')))
            if pdf_commande and beeline_commande and pdf_commande == beeline_commande:
                score += 3
                matched_fields.append("commande")
            elif pdf_commande and beeline_commande and fuzz.ratio(pdf_commande, beeline_commande) > threshold:
                score += 2
                matched_fields.append("commande_fuzzy")
            else:
                continue  # Pas de correspondance commande = pas de match
            
            # Correspondance montant net (avec tolérance)
            beeline_montant = beeline_row.get('montant_net_fournisseur', 0) or 0
            if pdf_total_net > 0 and beeline_montant > 0:
                diff_relative = abs(pdf_total_net - beeline_montant) / max(pdf_total_net, beeline_montant)
                if diff_relative < 0.05:  # 5% de tolérance
                    score += 2
                    matched_fields.append("montant_exact")
                elif diff_relative < 0.15:  # 15% de tolérance
                    score += 1
                    matched_fields.append("montant_proche")
            
            # Correspondance code rubrique
            beeline_rubrique = normalize_text(str(beeline_row.get('code_rubrique', '')))
            if pdf_rubrique and beeline_rubrique:
                if pdf_rubrique == beeline_rubrique:
                    score += 1
                    matched_fields.append("rubrique")
                elif fuzz.ratio(pdf_rubrique, beeline_rubrique) > threshold:
                    score += 0.5
                    matched_fields.append("rubrique_fuzzy")
            
            # Correspondance par similarité de texte globale
            pdf_text = f"{pdf_commande} {pdf_rubrique} {pdf_batch}"
            beeline_text = f"{beeline_commande} {beeline_rubrique} {normalize_text(str(beeline_row.get('collaborateur', '')))}"
            
            if fuzz.partial_ratio(pdf_text, beeline_text) > threshold:
                score += 0.5
                matched_fields.append("similarite_globale")
            
            if score > best_score:
                best_score = score
                best_match = beeline_row
                best_matched_fields = matched_fields
        
        return best_match, best_score, best_matched_fields
    
    def perform_smart_matching(self, tolerance: float = 0.05, fuzzy_threshold: int = 85) -> Dict:
        """Effectue un rapprochement intelligent hybride"""
        
        matched_pairs = []
        unmatched_pdf = []
        unmatched_beeline = []
        
        remaining_pdf = self.excel_pdf_data.copy()
        remaining_beeline = self.excel_beeline_data.copy()
        
        st.write("🔍 **RAPPROCHEMENT HYBRIDE INTELLIGENT**")
        
        # Analyser les numéros de commande
        pdf_commandes = set(item.get('numero_commande') for item in remaining_pdf if item.get('numero_commande'))
        beeline_commandes = set(item.get('numero_commande') for item in remaining_beeline if item.get('numero_commande'))
        commandes_communes = pdf_commandes.intersection(beeline_commandes)
        
        st.write(f"📊 **PDF** : {len(pdf_commandes)} commandes uniques : {sorted(list(pdf_commandes))[:5]}{'...' if len(pdf_commandes) > 5 else ''}")
        st.write(f"📋 **Beeline** : {len(beeline_commandes)} commandes uniques : {sorted(list(beeline_commandes))[:5]}{'...' if len(beeline_commandes) > 5 else ''}")
        st.write(f"🔗 **Commandes en commun** : {len(commandes_communes)} : {sorted(list(commandes_communes))}")
        
        # Stratégie 1 : Correspondance par numéro de commande + fuzzy matching
        if len(commandes_communes) > 0:
            st.write("✅ **Stratégie 1** : Correspondance par numéro de commande")
            self.matching_method = "COMMANDE_MATCHING"
            
            for pdf_row in remaining_pdf.copy():
                pdf_commande = pdf_row.get('numero_commande')
                
                if not pdf_commande or pdf_commande not in commandes_communes:
                    continue
                
                # Trouver tous les candidats Beeline avec la même commande
                beeline_candidates = [b for b in remaining_beeline if b.get('numero_commande') == pdf_commande]
                
                if beeline_candidates:
                    best_match, best_score, matched_fields = self.fuzzy_match_beeline_row(
                        pdf_row, beeline_candidates, fuzzy_threshold
                    )
                    
                    if best_match and best_score >= 3:  # Score minimum pour match valide
                        match_data = {
                            'pdf_data': pdf_row,
                            'beeline_data': best_match,
                            'match_type': 'INTELLIGENT',
                            'confidence': min(0.99, 0.6 + (best_score * 0.1)),
                            'fuzzy_score': best_score,
                            'matched_fields': matched_fields,
                            'match_method': 'commande_fuzzy'
                        }
                        
                        matched_pairs.append(match_data)
                        remaining_pdf.remove(pdf_row)
                        remaining_beeline.remove(best_match)
        
        # Stratégie 2 : Correspondance par ordre (fallback comme Copilot)
        else:
            st.write("🔄 **Stratégie 2** : Correspondance par ordre de fichiers (comme Copilot)")
            self.matching_method = "ORDER_MATCHING"
            
            # Grouper par fichier source
            pdf_by_file = {}
            for item in remaining_pdf:
                file_name = item['source_file']
                if file_name not in pdf_by_file:
                    pdf_by_file[file_name] = []
                pdf_by_file[file_name].append(item)
            
            beeline_by_file = {}
            for item in remaining_beeline:
                file_name = item['source_file']
                if file_name not in beeline_by_file:
                    beeline_by_file[file_name] = []
                beeline_by_file[file_name].append(item)
            
            pdf_files = list(pdf_by_file.keys())
            beeline_files = list(beeline_by_file.keys())
            
            # Correspondance par ordre d'upload
            nb_match = min(len(pdf_files), len(beeline_files))
            
            for i in range(nb_match):
                pdf_file_data = pdf_by_file[pdf_files[i]]
                beeline_file_data = beeline_by_file[beeline_files[i]]
                
                st.write(f"   🔗 Correspondance : {pdf_files[i]} ↔ {beeline_files[i]}")
                
                # Calculer le total net de chaque fichier pour validation
                pdf_total = sum(item.get('total_net', 0) or 0 for item in pdf_file_data)
                beeline_total = sum(item.get('montant_net_fournisseur', 0) or 0 for item in beeline_file_data)
                
                # Validation par total (comme Copilot)
                total_match = abs(pdf_total - beeline_total) < 0.02 if pdf_total > 0 and beeline_total > 0 else False
                
                # Matcher chaque ligne PDF avec les lignes Beeline du même fichier
                for pdf_row in pdf_file_data:
                    best_match, best_score, matched_fields = self.fuzzy_match_beeline_row(
                        pdf_row, beeline_file_data, fuzzy_threshold
                    )
                    
                    if best_match:
                        match_data = {
                            'pdf_data': pdf_row,
                            'beeline_data': best_match,
                            'match_type': 'ORDER_BASED',
                            'confidence': 0.8 if total_match else 0.6,
                            'fuzzy_score': best_score,
                            'matched_fields': matched_fields,
                            'match_method': 'order_fuzzy',
                            'total_validation': total_match
                        }
                        
                        matched_pairs.append(match_data)
                        if best_match in beeline_file_data:
                            beeline_file_data.remove(best_match)
                        if pdf_row in remaining_pdf:
                            remaining_pdf.remove(pdf_row)
                        if best_match in remaining_beeline:
                            remaining_beeline.remove(best_match)
        
        # Stocker les résultats
        self.matched_data = matched_pairs
        self.unmatched_pdf = remaining_pdf
        self.unmatched_beeline = remaining_beeline
        
        # Calculer les statistiques
        total_pdf = len(self.excel_pdf_data)
        total_beeline = len(self.excel_beeline_data)
        total_matched = len(matched_pairs)
        
        self.matching_stats = {
            'total_pdf_rows': total_pdf,
            'total_beeline_rows': total_beeline,
            'total_matched': total_matched,
            'unmatched_pdf': len(self.unmatched_pdf),
            'unmatched_beeline': len(self.unmatched_beeline),
            'match_rate_pdf': (total_matched / total_pdf * 100) if total_pdf > 0 else 0,
            'match_rate_beeline': (total_matched / total_beeline * 100) if total_beeline > 0 else 0,
            'intelligent_matches': len([m for m in matched_pairs if m['match_type'] == 'INTELLIGENT']),
            'order_matches': len([m for m in matched_pairs if m['match_type'] == 'ORDER_BASED']),
            'matching_method': self.matching_method
        }
        
        st.write(f"✅ **Résultat** : {total_matched} correspondances trouvées ({self.matching_method})")
        
        return self.matching_stats
    
    def create_consolidated_report(self) -> io.BytesIO:
        """Crée le rapport consolidé Excel"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # Feuille 1: Données consolidées
            consolidated_data = []
            for match in self.matched_data:
                pdf = match['pdf_data']
                beeline = match['beeline_data']
                
                consolidated_row = {
                    # Données PDF
                    'Source_PDF': pdf['source_file'],
                    'Numero_Facture': pdf['numero_facture'],
                    'Date_Facture': pdf['date_facture'],
                    'Destinataire': pdf['destinataire'],
                    'Batch_ID': pdf['batch_id'],
                    'Assignment_ID': pdf['assignment_id'],
                    'Total_Net_PDF': pdf['total_net'],
                    'Total_Brut_PDF': pdf['total_brut'],
                    
                    # Données Beeline
                    'Source_Beeline': beeline['source_file'],
                    'Collaborateur': beeline['collaborateur'],
                    'Code_Rubrique': beeline['code_rubrique'],
                    'Taux_Facturation': beeline['taux_facturation'],
                    'Unites': beeline['unites'],
                    'Montant_Brut_Beeline': beeline['montant_brut'],
                    'Montant_Net_Fournisseur': beeline['montant_net_fournisseur'],
                    'Projet': beeline['projet'],
                    'Centre_Cout': beeline['centre_cout'],
                    
                    # Données communes
                    'Numero_Commande': pdf['numero_commande'],
                    'Semaine_PDF': pdf['semaine_finissant_le'],
                    'Semaine_Beeline': beeline['semaine_finissant_le'],
                    
                    # Métadonnées de matching
                    'Type_Match': match['match_type'],
                    'Methode_Match': match['match_method'],
                    'Confiance': round(match['confidence'], 3),
                    'Score_Fuzzy': match['fuzzy_score'],
                    'Champs_Matches': ', '.join(match['matched_fields']),
                    'Validation_Total': match.get('total_validation', False),
                    'Ecart_Montant': abs((pdf['total_net'] or 0) - (beeline['montant_net_fournisseur'] or 0))
                }
                
                consolidated_data.append(consolidated_row)
            
            if consolidated_data:
                df_consolidated = pd.DataFrame(consolidated_data)
                df_consolidated.to_excel(writer, sheet_name='Donnees_Consolidees', index=False)
            
            # Feuille 2: Non-matchés PDF
            if self.unmatched_pdf:
                unmatched_pdf_df = pd.DataFrame(self.unmatched_pdf)
                unmatched_pdf_df.to_excel(writer, sheet_name='Non_Matches_PDF', index=False)
            
            # Feuille 3: Non-matchés Beeline
            if self.unmatched_beeline:
                unmatched_beeline_df = pd.DataFrame(self.unmatched_beeline)
                unmatched_beeline_df.to_excel(writer, sheet_name='Non_Matches_Beeline', index=False)
            
            # Feuille 4: Statistiques
            stats_data = [
                ['Métrique', 'Valeur'],
                ['Total lignes PDF', self.matching_stats['total_pdf_rows']],
                ['Total lignes Beeline', self.matching_stats['total_beeline_rows']],
                ['Total matches trouvés', self.matching_stats['total_matched']],
                ['Matches intelligents', self.matching_stats['intelligent_matches']],
                ['Matches par ordre', self.matching_stats['order_matches']],
                ['PDF non-matchés', self.matching_stats['unmatched_pdf']],
                ['Beeline non-matchés', self.matching_stats['unmatched_beeline']],
                ['Taux de match PDF (%)', round(self.matching_stats['match_rate_pdf'], 2)],
                ['Taux de match Beeline (%)', round(self.matching_stats['match_rate_beeline'], 2)],
                ['Méthode utilisée', self.matching_stats['matching_method']]
            ]
            
            df_stats = pd.DataFrame(stats_data[1:], columns=stats_data[0])
            df_stats.to_excel(writer, sheet_name='Statistiques', index=False)
            
            # Feuille 5: Synthèse par collaborateur
            if consolidated_data:
                collaborateur_synthesis = {}
                for row in consolidated_data:
                    collab = row['Collaborateur']
                    if collab not in collaborateur_synthesis:
                        collaborateur_synthesis[collab] = {
                            'Collaborateur': collab,
                            'Nb_Matches': 0,
                            'Total_Net_PDF': 0,
                            'Total_Net_Beeline': 0,
                            'Commandes': set(),
                            'Rubriques': set(),
                            'Score_Moyen': 0,
                            'Confiance_Moyenne': 0
                        }
                    
                    synthesis = collaborateur_synthesis[collab]
                    synthesis['Nb_Matches'] += 1
                    synthesis['Total_Net_PDF'] += row['Total_Net_PDF'] or 0
                    synthesis['Total_Net_Beeline'] += row['Montant_Net_Fournisseur'] or 0
                    synthesis['Commandes'].add(row['Numero_Commande'])
                    synthesis['Rubriques'].add(row['Code_Rubrique'])
                    synthesis['Score_Moyen'] += row['Score_Fuzzy']
                    synthesis['Confiance_Moyenne'] += row['Confiance']
                
                synthesis_export = []
                for synthesis in collaborateur_synthesis.values():
                    synthesis_export.append({
                        'Collaborateur': synthesis['Collaborateur'],
                        'Nb_Matches': synthesis['Nb_Matches'],
                        'Total_Net_PDF': synthesis['Total_Net_PDF'],
                        'Total_Net_Beeline': synthesis['Total_Net_Beeline'],
                        'Ecart_Total': synthesis['Total_Net_PDF'] - synthesis['Total_Net_Beeline'],
                        'Score_Fuzzy_Moyen': round(synthesis['Score_Moyen'] / synthesis['Nb_Matches'], 2),
                        'Confiance_Moyenne': round(synthesis['Confiance_Moyenne'] / synthesis['Nb_Matches'], 3),
                        'Nb_Commandes': len(synthesis['Commandes']),
                        'Commandes': ', '.join(sorted(synthesis['Commandes'])),
                        'Rubriques': ', '.join(sorted(filter(None, synthesis['Rubriques'])))
                    })
                
                if synthesis_export:
                    df_synthesis = pd.DataFrame(synthesis_export)
                    df_synthesis.to_excel(writer, sheet_name='Synthese_Collaborateurs', index=False)
        
        output.seek(0)
        return output


def main():
    st.title("🔗 Excel Matcher Hybride Intelligent")
    st.markdown("### Rapprochement Excel PDF ↔ Excel Beeline avec IA adaptative")
    
    # Sidebar
    st.sidebar.header("⚙️ Paramètres de matching")
    
    tolerance = st.sidebar.slider(
        "Tolérance sur les montants (%)", 
        min_value=1, 
        max_value=20, 
        value=5, 
        help="Tolérance acceptée pour considérer que deux montants correspondent"
    ) / 100
    
    fuzzy_threshold = st.sidebar.slider(
        "Seuil de similarité fuzzy",
        min_value=70,
        max_value=100,
        value=85,
        help="Plus élevé = plus strict (comme Copilot)"
    )
    
    st.sidebar.header("🧠 Intelligence adaptive")
    st.sidebar.markdown("""
    **L'app choisit automatiquement :**
    1. **Correspondance par N° commande** si possible
    2. **Correspondance par ordre** (comme Copilot) sinon
    3. **Matching fuzzy** dans tous les cas
    """)
    
    st.sidebar.header("📋 Instructions")
    st.sidebar.markdown("""
    1. **Uploadez** vos Excel PDF (App 1)
    2. **Uploadez** vos Excel Beeline  
    3. **Ajustez** les paramètres si besoin
    4. **Lancez** le rapprochement intelligent
    5. **Téléchargez** le rapport consolidé
    """)
    
    # Section 1: Upload Excel PDF
    st.header("📊 1. Fichiers Excel PDF (issus de l'App 1)")
    
    uploaded_excel_pdf = st.file_uploader(
        "Sélectionnez vos fichiers Excel extraits des PDFs",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Fichiers Excel générés par l'App 1 (extraction PDF)",
        key="excel_pdf"
    )
    
    if uploaded_excel_pdf:
        st.success(f"✅ {len(uploaded_excel_pdf)} fichier(s) Excel PDF sélectionné(s)")
        
        with st.expander("📁 Fichiers Excel PDF sélectionnés", expanded=False):
            for i, file in enumerate(uploaded_excel_pdf, 1):
                st.write(f"{i}. {file.name} ({file.size / 1024:.1f} KB)")
    
    # Section 2: Upload Excel Beeline
    st.header("📋 2. Fichiers Excel Beeline")
    
    uploaded_excel_beeline = st.file_uploader(
        "Sélectionnez vos fichiers Excel Beeline",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Fichiers Supplier Payment Register de Beeline",
        key="excel_beeline"
    )
    
    if uploaded_excel_beeline:
        st.success(f"✅ {len(uploaded_excel_beeline)} fichier(s) Excel Beeline sélectionné(s)")
        
        with st.expander("📁 Fichiers Excel Beeline sélectionnés", expanded=False):
            for i, file in enumerate(uploaded_excel_beeline, 1):
                st.write(f"{i}. {file.name} ({file.size / 1024:.1f} KB)")
    
    # Section 3: Lancement du matching intelligent
    if uploaded_excel_pdf and uploaded_excel_beeline:
        st.header("🧠 3. Rapprochement intelligent")
        
        if st.button("🚀 Lancer le rapprochement hybride", type="primary"):
            with st.spinner("Analyse intelligente en cours..."):
                
                matcher = HybridExcelMatcher()
                
                # Phase 1: Chargement des données
                st.subheader("📊 Chargement et analyse des données")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**📊 Chargement Excel PDF...**")
                    pdf_data = matcher.load_excel_pdf_files(uploaded_excel_pdf)
                    st.success(f"✅ {len(pdf_data)} lignes Excel PDF chargées")
                
                with col2:
                    st.write("**📋 Chargement Excel Beeline...**")
                    beeline_data = matcher.load_excel_beeline_files(uploaded_excel_beeline)
                    st.success(f"✅ {len(beeline_data)} lignes Excel Beeline chargées")
                
                if len(pdf_data) == 0:
                    st.error("❌ Aucune donnée valide trouvée dans les fichiers Excel PDF")
                    return
                
                if len(beeline_data) == 0:
                    st.error("❌ Aucune donnée valide trouvée dans les fichiers Excel Beeline")
                    return
                
                # Phase 2: Rapprochement intelligent
                st.subheader("🧠 Rapprochement hybride intelligent")
                
                matching_stats = matcher.perform_smart_matching(tolerance, fuzzy_threshold)
                
                # Phase 3: Affichage des résultats
                st.header("📊 Résultats du rapprochement")
                
                # Métriques principales
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric("📊 Lignes PDF", matching_stats['total_pdf_rows'])
                
                with col2:
                    st.metric("📋 Lignes Beeline", matching_stats['total_beeline_rows'])
                
                with col3:
                    st.metric("🔗 Matches trouvés", matching_stats['total_matched'])
                
                with col4:
                    st.metric("✅ Taux match PDF", f"{matching_stats['match_rate_pdf']:.1f}%")
                
                with col5:
                    st.metric("🧠 Méthode utilisée", matching_stats['matching_method'])
                
                # Détail des types de matches
                st.subheader("🎯 Détail des correspondances")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("🎯 Matches intelligents", matching_stats['intelligent_matches'])
                
                with col2:
                    st.metric("🔄 Matches par ordre", matching_stats['order_matches'])
                
                with col3:
                    st.metric("❌ PDF non-matchés", matching_stats['unmatched_pdf'])
                
                with col4:
                    st.metric("❌ Beeline non-matchés", matching_stats['unmatched_beeline'])
                
                # Aperçu des matches avec détails du fuzzy matching
                if matcher.matched_data:
                    st.subheader("👥 Aperçu des correspondances trouvées")
                    
                    matches_display = []
                    for match in matcher.matched_data[:10]:
                        pdf = match['pdf_data']
                        beeline = match['beeline_data']
                        
                        matches_display.append({
                            'Type': match['match_type'],
                            'Méthode': match['match_method'],
                            'Score Fuzzy': match['fuzzy_score'],
                            'Confiance': f"{match['confidence']:.2f}",
                            'Champs matchés': ', '.join(match['matched_fields']),
                            'N° Facture': pdf['numero_facture'],
                            'N° Commande': pdf['numero_commande'],
                            'Collaborateur': beeline['collaborateur'],
                            'Net PDF (€)': f"{pdf['total_net']:,.2f}" if pdf['total_net'] else "0,00",
                            'Net Beeline (€)': f"{beeline['montant_net_fournisseur']:,.2f}" if beeline['montant_net_fournisseur'] else "0,00",
                            'Écart (€)': f"{abs((pdf['total_net'] or 0) - (beeline['montant_net_fournisseur'] or 0)):,.2f}",
                        })
                    
                    df_matches = pd.DataFrame(matches_display)
                    st.dataframe(df_matches, use_container_width=True)
                    
                    if len(matcher.matched_data) > 10:
                        st.info(f"ℹ️ Affichage des 10 premiers matches sur {len(matcher.matched_data)} au total")
                
                # Analyse de la qualité du matching
                if matcher.matched_data:
                    st.subheader("📈 Analyse de la qualité du matching")
                    
                    # Statistiques de confiance
                    confidences = [m['confidence'] for m in matcher.matched_data]
                    fuzzy_scores = [m['fuzzy_score'] for m in matcher.matched_data]
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        if confidences:
                            avg_confidence = sum(confidences) / len(confidences)
                            st.metric("🎯 Confiance moyenne", f"{avg_confidence:.3f}")
                        else:
                            st.metric("🎯 Confiance moyenne", "N/A")
                    
                    with col2:
                        if fuzzy_scores:
                            avg_fuzzy = sum(fuzzy_scores) / len(fuzzy_scores)
                            st.metric("🔍 Score fuzzy moyen", f"{avg_fuzzy:.1f}")
                        else:
                            st.metric("🔍 Score fuzzy moyen", "N/A")
                    
                    with col3:
                        if confidences:
                            high_confidence = len([c for c in confidences if c > 0.8])
                            st.metric("⭐ Matches haute confiance", f"{high_confidence}")
                        else:
                            st.metric("⭐ Matches haute confiance", "0")
                
                # Analyse par collaborateur
                if matcher.matched_data:
                    st.subheader("👥 Analyse par collaborateur")
                    
                    collaborateurs = {}
                    for match in matcher.matched_data:
                        collab = match['beeline_data']['collaborateur']
                        if collab not in collaborateurs:
                            collaborateurs[collab] = {
                                'nb_matches': 0,
                                'total_net_pdf': 0,
                                'total_net_beeline': 0,
                                'commandes': set(),
                                'score_moyen': 0,
                                'confiance_moyenne': 0
                            }
                        
                        collaborateurs[collab]['nb_matches'] += 1
                        collaborateurs[collab]['total_net_pdf'] += match['pdf_data']['total_net'] or 0
                        collaborateurs[collab]['total_net_beeline'] += match['beeline_data']['montant_net_fournisseur'] or 0
                        collaborateurs[collab]['commandes'].add(match['pdf_data']['numero_commande'])
                        collaborateurs[collab]['score_moyen'] += match['fuzzy_score']
                        collaborateurs[collab]['confiance_moyenne'] += match['confidence']
                    
                    collab_display = []
                    for name, stats in sorted(collaborateurs.items()):
                        collab_display.append({
                            'Collaborateur': name,
                            'Nb Matches': stats['nb_matches'],
                            'Nb Commandes': len(stats['commandes']),
                            'Score Fuzzy Moyen': round(stats['score_moyen'] / stats['nb_matches'], 1),
                            'Confiance Moyenne': round(stats['confiance_moyenne'] / stats['nb_matches'], 3),
                            'Total Net PDF (€)': f"{stats['total_net_pdf']:,.2f}",
                            'Total Net Beeline (€)': f"{stats['total_net_beeline']:,.2f}",
                            'Écart (€)': f"{abs(stats['total_net_pdf'] - stats['total_net_beeline']):,.2f}"
                        })
                    
                    df_collab = pd.DataFrame(collab_display)
                    st.dataframe(df_collab, use_container_width=True)
                
                # Détail des non-matchés
                with st.expander("❌ Données non rapprochées", expanded=False):
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("**📊 PDF non-matchés:**")
                        if matcher.unmatched_pdf:
                            unmatched_pdf_display = []
                            for item in matcher.unmatched_pdf[:5]:
                                unmatched_pdf_display.append({
                                    'N° Facture': item['numero_facture'],
                                    'N° Commande': item['numero_commande'],
                                    'Total Net (€)': f"{item['total_net']:,.2f}" if item['total_net'] else "0,00",
                                    'Fichier': item['source_file']
                                })
                            
                            df_unmatched_pdf = pd.DataFrame(unmatched_pdf_display)
                            st.dataframe(df_unmatched_pdf, use_container_width=True)
                            
                            if len(matcher.unmatched_pdf) > 5:
                                st.info(f"ℹ️ +{len(matcher.unmatched_pdf) - 5} autres...")
                        else:
                            st.success("✅ Tous les PDF ont été rapprochés")
                    
                    with col2:
                        st.markdown("**📋 Beeline non-matchés:**")
                        if matcher.unmatched_beeline:
                            unmatched_beeline_display = []
                            for item in matcher.unmatched_beeline[:5]:
                                unmatched_beeline_display.append({
                                    'Collaborateur': item['collaborateur'],
                                    'N° Commande': item['numero_commande'],
                                    'Net Fournisseur (€)': f"{item['montant_net_fournisseur']:,.2f}" if item['montant_net_fournisseur'] else "0,00",
                                    'Fichier': item['source_file']
                                })
                            
                            df_unmatched_beeline = pd.DataFrame(unmatched_beeline_display)
                            st.dataframe(df_unmatched_beeline, use_container_width=True)
                            
                            if len(matcher.unmatched_beeline) > 5:
                                st.info(f"ℹ️ +{len(matcher.unmatched_beeline) - 5} autres...")
                        else:
                            st.success("✅ Tous les Beeline ont été rapprochés")
                
                # Génération du rapport final
                st.header("💾 Export du rapport consolidé")
                
                with st.spinner("Génération du rapport Excel..."):
                    excel_report = matcher.create_consolidated_report()
                
                # Informations sur le rapport
                st.subheader("📋 Contenu du rapport Excel")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("""
                    **🎯 Feuilles incluses :**
                    - **Donnees_Consolidees** : Correspondances avec métadonnées fuzzy
                    - **Non_Matches_PDF** : Données PDF non rapprochées
                    - **Non_Matches_Beeline** : Données Beeline non rapprochées
                    """)
                
                with col2:
                    st.markdown("""
                    **📊 Analyses incluses :**
                    - **Statistiques** : Métriques de performance hybride
                    - **Synthese_Collaborateurs** : Analyse détaillée par intérimaire
                    - **Scores fuzzy et confiance** : Qualité des rapprochements
                    """)
                
                # Bouton de téléchargement
                filename = f"Rapport_Hybride_PDF_Beeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                st.download_button(
                    label="📊 Télécharger le rapport hybride consolidé",
                    data=excel_report.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                # Résumé final
                st.success("🎉 Rapprochement hybride terminé avec succès !")
                
                with st.expander("📈 Résumé détaillé", expanded=False):
                    st.markdown(f"""
                    **🧠 Méthode intelligente utilisée :**
                    - {matching_stats['matching_method']}
                    
                    **📊 Données traitées :**
                    - {matching_stats['total_pdf_rows']} lignes Excel PDF
                    - {matching_stats['total_beeline_rows']} lignes Excel Beeline
                    
                    **🔗 Correspondances :**
                    - {matching_stats['intelligent_matches']} matches intelligents (par N° commande)
                    - {matching_stats['order_matches']} matches par ordre (comme Copilot)
                    - {matching_stats['total_matched']} total matches
                    
                    **📈 Performance :**
                    - {matching_stats['match_rate_pdf']:.1f}% des PDF rapprochés
                    - {matching_stats['match_rate_beeline']:.1f}% des Beeline rapprochés
                    
                    **🎯 Qualité :**
                    - Matching fuzzy avec seuil {fuzzy_threshold}%
                    - Confiance moyenne: {(sum(m['confidence'] for m in matcher.matched_data) / len(matcher.matched_data)):.3f if len(matcher.matched_data) > 0 else 'N/A'}
                    - Enrichissement avec noms des intérimaires et validation croisée
                    """)
    
    else:
        st.info("👆 Commencez par uploader vos fichiers Excel PDF et Beeline")
    
    # Footer
    st.markdown("---")
    st.markdown("**Excel Matcher Hybride** - Version 1.0 | IA adaptative pour rapprochement PDF ↔ Beeline")


if __name__ == "__main__":
    main()
