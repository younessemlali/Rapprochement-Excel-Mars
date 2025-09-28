import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta
import re
from typing import Dict, List, Tuple, Optional
import logging
from difflib import SequenceMatcher

# Configuration de la page
st.set_page_config(
    page_title="Excel Matcher Beeline",
    page_icon="🔗",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ExcelMatcher:
    """Classe pour le rapprochement des fichiers Excel PDF et Beeline"""
    
    def __init__(self):
        self.excel_pdf_data = []
        self.excel_beeline_data = []
        self.matched_data = []
        self.unmatched_pdf = []
        self.unmatched_beeline = []
        self.matching_stats = {}
    
    def load_excel_pdf_files(self, uploaded_files) -> List[Dict]:
        """Charge et traite les fichiers Excel issus de l'App 1 (PDF)"""
        all_data = []
        
        for uploaded_file in uploaded_files:
            st.write(f"📊 Traitement Excel PDF: {uploaded_file.name}")
            
            try:
                # Lire toutes les feuilles du fichier Excel
                excel_data = pd.read_excel(uploaded_file, sheet_name=None)
                
                st.write(f"   📋 Feuilles trouvées: {list(excel_data.keys())}")
                
                # Définir l'ordre de priorité des feuilles à traiter
                priority_sheets = ['Donnees_Analyse', 'Résumé_Factures', 'Detail_Lignes', 'Analyse_Rubriques']
                
                sheets_to_process = []
                
                # Ajouter d'abord les feuilles prioritaires
                for sheet_name in priority_sheets:
                    if sheet_name in excel_data:
                        sheets_to_process.append(sheet_name)
                
                # Ajouter toutes les autres feuilles
                for sheet_name in excel_data.keys():
                    if sheet_name not in sheets_to_process:
                        sheets_to_process.append(sheet_name)
                
                st.write(f"   🔄 Feuilles à traiter: {sheets_to_process}")
                
                total_rows_added = 0
                
                # TRAITER TOUTES LES FEUILLES
                for sheet_name in sheets_to_process:
                    df_to_use = excel_data[sheet_name]
                    
                    if df_to_use is None or len(df_to_use) == 0:
                        st.write(f"   ⚠️ Feuille '{sheet_name}' vide, ignorée")
                        continue
                    
                    st.write(f"   📊 Traitement feuille '{sheet_name}' ({len(df_to_use)} lignes)")
                    
                    # Nettoyer et standardiser les données
                    df_cleaned = self.clean_pdf_excel_data(df_to_use)
                    
                    st.write(f"     📋 Colonnes: {list(df_cleaned.columns)[:5]}{'...' if len(df_cleaned.columns) > 5 else ''}")
                    
                    # Debug: Afficher les premières lignes si contient des colonnes intéressantes
                    has_useful_columns = any(col in df_cleaned.columns for col in ['Numero_Facture', 'Numero_Commande', 'N° Facture', 'N° Commande'])
                    
                    if has_useful_columns and len(df_cleaned) > 0:
                        st.write(f"     🔍 Échantillon première ligne:")
                        relevant_cols = ['Numero_Facture', 'Numero_Commande', 'N° Facture', 'N° Commande', 'Total_Net_EUR', 'Total_Net']
                        for col in relevant_cols:
                            if col in df_cleaned.columns:
                                st.write(f"       - {col}: {df_cleaned.iloc[0][col]}")
                    
                    rows_added_sheet = 0
                    for _, row in df_cleaned.iterrows():
                        # Chercher le vrai numéro de commande dans le nom du fichier
                        true_commande = self.extract_commande_from_filename(uploaded_file.name)
                        
                        data_row = {
                            'source_file': uploaded_file.name,
                            'source_sheet': sheet_name,
                            'numero_facture': row.get('Numero_Facture') or row.get('N° Facture'),
                            'numero_commande': true_commande or row.get('Numero_Commande') or row.get('N° Commande'),
                            'date_facture': row.get('Date_Facture') or row.get('Date'),
                            'semaine_finissant_le': row.get('Semaine_Finissant_Le') or self.extract_week_from_date(row.get('Date_Facture')) or self.extract_week_from_date(row.get('Date_Periode')),
                            'destinataire': row.get('Destinataire'),
                            'batch_id': row.get('Batch_ID'),
                            'assignment_id': row.get('Assignment_ID'),
                            'total_net': self.safe_float(row.get('Total_Net_EUR') or row.get('Total_Net') or row.get('Montant_Net')),
                            'total_tva': self.safe_float(row.get('Total_TVA_EUR') or row.get('Total_TVA') or row.get('Montant_TVA')),
                            'total_brut': self.safe_float(row.get('Total_Brut_EUR') or row.get('Total_Brut') or row.get('Montant_Brut')),
                            'type_donnees': 'PDF_EXTRACT'
                        }
                        
                        # Ajouter seulement si on a au minimum un numéro de commande
                        if data_row['numero_commande']:
                            all_data.append(data_row)
                            rows_added_sheet += 1
                    
                    st.write(f"     ✅ {rows_added_sheet} lignes valides ajoutées de '{sheet_name}'")
                    total_rows_added += rows_added_sheet
                
                st.write(f"   🎯 TOTAL: {total_rows_added} lignes ajoutées du fichier {uploaded_file.name}")
                
            except Exception as e:
                st.error(f"❌ Erreur lors du traitement de {uploaded_file.name}: {e}")
                logger.error(f"Erreur Excel PDF {uploaded_file.name}: {e}")
        
        self.excel_pdf_data = all_data
        return all_data
    
    def load_excel_beeline_files(self, uploaded_files) -> List[Dict]:
        """Charge et traite les fichiers Excel Beeline"""
        all_data = []
        
        for uploaded_file in uploaded_files:
            st.write(f"📋 Traitement Excel Beeline: {uploaded_file.name}")
            
            try:
                # Lire le fichier Excel (généralement une seule feuille)
                df = pd.read_excel(uploaded_file)
                
                # Nettoyer et standardiser les données
                df_cleaned = self.clean_beeline_excel_data(df)
                
                st.write(f"   📋 Colonnes trouvées: {list(df_cleaned.columns)}")
                st.write(f"   📊 {len(df_cleaned)} lignes dans le fichier")
                
                # Debug: Afficher les premières lignes
                if len(df_cleaned) > 0:
                    st.write(f"   🔍 Échantillon première ligne:")
                    for col in df_cleaned.columns[:8]:  # Premières 8 colonnes
                        if col in df_cleaned.columns:
                            st.write(f"     - {col}: {df_cleaned.iloc[0][col]}")
                
                rows_added = 0
                for _, row in df_cleaned.iterrows():
                    data_row = {
                        'source_file': uploaded_file.name,
                        'collaborateur': row.get('Collaborateur'),
                        'numero_commande': row.get('N° commande') or row.get('Numero_Commande'),
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
                    
                    # Ajouter seulement si on a au minimum un numéro de commande et un collaborateur
                    if data_row['numero_commande'] and data_row['collaborateur']:
                        all_data.append(data_row)
                        rows_added += 1
                
                st.write(f"   ✅ {rows_added} lignes valides ajoutées (avec N° commande + collaborateur)")
                
            except Exception as e:
                st.error(f"❌ Erreur lors du traitement de {uploaded_file.name}: {e}")
                logger.error(f"Erreur Excel Beeline {uploaded_file.name}: {e}")
        
        self.excel_beeline_data = all_data
        return all_data
    
    def clean_pdf_excel_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Nettoie les données Excel PDF"""
        # Supprimer les lignes vides
        df = df.dropna(how='all')
        
        # Standardiser les noms de colonnes
        df.columns = df.columns.str.strip()
        
        return df
    
    def clean_beeline_excel_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Nettoie les données Excel Beeline"""
        # Supprimer les lignes vides
        df = df.dropna(how='all')
        
        # Standardiser les noms de colonnes
        df.columns = df.columns.str.strip()
        
        return df
    
    def safe_float(self, value) -> Optional[float]:
        """Convertit une valeur en float de manière sécurisée"""
        if value is None or pd.isna(value):
            return None
        
        try:
            if isinstance(value, str):
                # Nettoyer la chaîne (espaces, symboles)
                value = re.sub(r'[^\d\.,\-]', '', str(value))
                if not value:
                    return None
                
                # Gérer les formats français (virgule = décimales)
                if ',' in value and '.' not in value:
                    value = value.replace(',', '.')
                elif ',' in value and '.' in value:
                    # Format 1,234.56 - supprimer la virgule
                    value = value.replace(',', '')
            
            return float(value)
        except (ValueError, TypeError):
            return None
    
    def extract_commande_from_filename(self, filename: str) -> Optional[str]:
        """Extrait le numéro de commande depuis le nom du fichier PDF"""
        try:
            # Pattern pour des noms comme: 123_4949S0001_1182_0015030425_5600025054_MARSFR_11032025.pdf
            # Le numéro de commande est souvent en 5ème position
            parts = filename.split('_')
            
            # Chercher un numéro qui ressemble à une commande (10 chiffres commençant par 56)
            for part in parts:
                if len(part) == 10 and part.startswith('56') and part.isdigit():
                    return part
            
            # Si pas trouvé, chercher d'autres patterns
            commande_patterns = [
                r'(56\d{8})',  # 56 suivi de 8 chiffres
                r'(\d{10})',   # 10 chiffres
            ]
            
            for pattern in commande_patterns:
                matches = re.findall(pattern, filename)
                if matches:
                    return matches[0]
            
            return None
            
        except Exception as e:
            logger.warning(f"Impossible d'extraire commande de {filename}: {e}")
            return None
    
    def extract_week_from_date(self, date_str) -> Optional[str]:
        """Extrait une semaine approximative depuis une date"""
        if not date_str:
            return None
        
        try:
            # Formats de date possibles
            date_formats = ['%Y/%m/%d', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y']
            
            date_obj = None
            for fmt in date_formats:
                try:
                    date_obj = datetime.strptime(str(date_str), fmt)
                    break
                except ValueError:
                    continue
            
            if date_obj:
                # Calculer le vendredi de cette semaine
                days_ahead = 4 - date_obj.weekday()  # 4 = vendredi
                if days_ahead < 0:  # Si on est après vendredi
                    days_ahead += 7
                
                friday = date_obj + timedelta(days=days_ahead)
                return friday.strftime('%d/%m/%Y')
            
        except Exception as e:
            logger.warning(f"Impossible d'extraire la semaine de {date_str}: {e}")
        
        return None
    
    def normalize_week_format(self, week_str) -> Optional[str]:
        """Normalise le format des semaines"""
        if not week_str:
            return None
        
        try:
            # Différents formats possibles
            week_str = str(week_str).strip()
            
            # Format DD/MM/YYYY ou DD-MM-YYYY
            if re.match(r'\d{1,2}[/-]\d{1,2}[/-]\d{4}', week_str):
                date_formats = ['%d/%m/%Y', '%d-%m-%Y']
                for fmt in date_formats:
                    try:
                        date_obj = datetime.strptime(week_str, fmt)
                        return date_obj.strftime('%d/%m/%Y')
                    except ValueError:
                        continue
            
            # Format YYYY/MM/DD ou YYYY-MM-DD
            elif re.match(r'\d{4}[/-]\d{1,2}[/-]\d{1,2}', week_str):
                date_formats = ['%Y/%m/%d', '%Y-%m-%d']
                for fmt in date_formats:
                    try:
                        date_obj = datetime.strptime(week_str, fmt)
                        return date_obj.strftime('%d/%m/%Y')
                    except ValueError:
                        continue
            
            return week_str  # Retourner tel quel si pas de conversion possible
            
        except Exception:
            return week_str
    
    def normalize_commande_number(self, commande_str) -> Optional[str]:
        """Normalise les numéros de commande"""
        if not commande_str:
            return None
        
        # Convertir en string et nettoyer
        commande_clean = str(commande_str).strip()
        
        # Supprimer les caractères non numériques
        commande_clean = re.sub(r'[^\d]', '', commande_clean)
        
        return commande_clean if commande_clean else None
    
    def calculate_amount_similarity(self, amount1: float, amount2: float, tolerance: float = 0.05) -> Tuple[bool, float]:
        """Calcule la similarité entre deux montants avec tolérance"""
        if amount1 is None or amount2 is None:
            return False, 0.0
        
        if amount1 == 0 and amount2 == 0:
            return True, 1.0
        
        if amount1 == 0 or amount2 == 0:
            return False, 0.0
        
        # Calculer la différence relative
        diff = abs(amount1 - amount2) / max(abs(amount1), abs(amount2))
        similarity = 1.0 - diff
        
        is_similar = diff <= tolerance
        
        return is_similar, similarity
    
    def perform_matching(self, tolerance: float = 0.05) -> Dict:
        """Effectue le rapprochement entre les données PDF et Beeline"""
        
        matched_pairs = []
        unmatched_pdf = []
        unmatched_beeline = []
        
        # Créer des copies pour le matching
        remaining_pdf = self.excel_pdf_data.copy()
        remaining_beeline = self.excel_beeline_data.copy()
        
        # DEBUT DEBUG - Afficher des échantillons de données
        st.write("🔍 **DEBUG - Échantillons de données pour diagnostic :**")
        
        if remaining_pdf:
            st.write("📊 **Échantillon données PDF :**")
            pdf_sample = remaining_pdf[0]
            st.write(f"- N° Commande: '{pdf_sample.get('numero_commande')}' (type: {type(pdf_sample.get('numero_commande'))})")
            st.write(f"- Semaine: '{pdf_sample.get('semaine_finissant_le')}' (type: {type(pdf_sample.get('semaine_finissant_le'))})")
            st.write(f"- Total Net: {pdf_sample.get('total_net')} (type: {type(pdf_sample.get('total_net'))})")
            
            # Normalisation de test
            pdf_commande_norm = self.normalize_commande_number(pdf_sample.get('numero_commande'))
            pdf_semaine_norm = self.normalize_week_format(pdf_sample.get('semaine_finissant_le'))
            st.write(f"- N° Commande normalisé: '{pdf_commande_norm}'")
            st.write(f"- Semaine normalisée: '{pdf_semaine_norm}'")
        
        if remaining_beeline:
            st.write("📋 **Échantillon données Beeline :**")
            beeline_sample = remaining_beeline[0]
            st.write(f"- N° Commande: '{beeline_sample.get('numero_commande')}' (type: {type(beeline_sample.get('numero_commande'))})")
            st.write(f"- Semaine: '{beeline_sample.get('semaine_finissant_le')}' (type: {type(beeline_sample.get('semaine_finissant_le'))})")
            st.write(f"- Montant Net Fournisseur: {beeline_sample.get('montant_net_fournisseur')} (type: {type(beeline_sample.get('montant_net_fournisseur'))})")
            st.write(f"- Collaborateur: '{beeline_sample.get('collaborateur')}'")
            
            # Normalisation de test
            beeline_commande_norm = self.normalize_commande_number(beeline_sample.get('numero_commande'))
            beeline_semaine_norm = self.normalize_week_format(beeline_sample.get('semaine_finissant_le'))
            st.write(f"- N° Commande normalisé: '{beeline_commande_norm}'")
            st.write(f"- Semaine normalisée: '{beeline_semaine_norm}'")
        
        # Afficher tous les numéros de commande uniques pour comparaison
        pdf_commandes = set(self.normalize_commande_number(item.get('numero_commande')) for item in remaining_pdf if item.get('numero_commande'))
        beeline_commandes = set(self.normalize_commande_number(item.get('numero_commande')) for item in remaining_beeline if item.get('numero_commande'))
        
        st.write(f"📊 **N° Commandes uniques PDF ({len(pdf_commandes)}):** {sorted(list(pdf_commandes))[:10]}{'...' if len(pdf_commandes) > 10 else ''}")
        st.write(f"📋 **N° Commandes uniques Beeline ({len(beeline_commandes)}):** {sorted(list(beeline_commandes))[:10]}{'...' if len(beeline_commandes) > 10 else ''}")
        
        # Chercher les intersections
        commandes_communes = pdf_commandes.intersection(beeline_commandes)
        st.write(f"🔗 **N° Commandes en commun ({len(commandes_communes)}):** {sorted(list(commandes_communes))}")
        
        if len(commandes_communes) == 0:
            st.error("❌ **PROBLÈME IDENTIFIÉ** : Aucun numéro de commande en commun entre PDF et Beeline !")
            st.write("🔍 **Vérifications à faire :**")
            st.write("1. Les fichiers Excel PDF contiennent-ils bien les bonnes colonnes ?")
            st.write("2. Les fichiers Beeline sont-ils au bon format ?")
            st.write("3. Les numéros de commande correspondent-ils entre les deux sources ?")
        # FIN DEBUG
        
        # Phase 1: Matching exact par N° commande + Semaine
        st.write("🔍 Phase 1: Matching exact par N° commande + Semaine...")
        
        matches_phase1 = 0
        for pdf_row in remaining_pdf.copy():
            pdf_commande = self.normalize_commande_number(pdf_row['numero_commande'])
            pdf_semaine = self.normalize_week_format(pdf_row['semaine_finissant_le'])
            
            if not pdf_commande:
                continue
            
            for beeline_row in remaining_beeline.copy():
                beeline_commande = self.normalize_commande_number(beeline_row['numero_commande'])
                beeline_semaine = self.normalize_week_format(beeline_row['semaine_finissant_le'])
                
                # Match exact sur commande
                if pdf_commande == beeline_commande:
                    # Vérifier la semaine si disponible
                    week_match = True
                    if pdf_semaine and beeline_semaine:
                        week_match = pdf_semaine == beeline_semaine
                    
                    if week_match:
                        # Vérifier la similarité des montants
                        amount_similar, similarity = self.calculate_amount_similarity(
                            pdf_row['total_net'], 
                            beeline_row['montant_net_fournisseur'], 
                            tolerance
                        )
                        
                        match_data = {
                            'pdf_data': pdf_row,
                            'beeline_data': beeline_row,
                            'match_type': 'EXACT',
                            'confidence': 0.95 if amount_similar else 0.75,
                            'amount_similarity': similarity,
                            'amount_match': amount_similar,
                            'week_match': week_match,
                            'commande_match': True
                        }
                        
                        matched_pairs.append(match_data)
                        remaining_pdf.remove(pdf_row)
                        remaining_beeline.remove(beeline_row)
                        matches_phase1 += 1
                        break
        
        st.write(f"✅ Phase 1 terminée: {matches_phase1} matches exacts trouvés")
        
        # Phase 2: Matching partiel par N° commande seulement
        st.write("🔍 Phase 2: Matching partiel par N° commande...")
        
        matches_phase2 = 0
        for pdf_row in remaining_pdf.copy():
            pdf_commande = self.normalize_commande_number(pdf_row['numero_commande'])
            
            if not pdf_commande:
                continue
            
            best_match = None
            best_similarity = 0
            
            for beeline_row in remaining_beeline:
                beeline_commande = self.normalize_commande_number(beeline_row['numero_commande'])
                
                if pdf_commande == beeline_commande:
                    # Calculer similarité globale
                    amount_similar, amount_sim = self.calculate_amount_similarity(
                        pdf_row['total_net'], 
                        beeline_row['montant_net_fournisseur'], 
                        tolerance * 2  # Tolérance plus large
                    )
                    
                    total_similarity = amount_sim * 0.8 + 0.2  # Base 20% pour match commande
                    
                    if total_similarity > best_similarity and total_similarity > 0.4:
                        best_similarity = total_similarity
                        best_match = beeline_row
            
            if best_match:
                match_data = {
                    'pdf_data': pdf_row,
                    'beeline_data': best_match,
                    'match_type': 'PARTIAL',
                    'confidence': best_similarity,
                    'amount_similarity': self.calculate_amount_similarity(pdf_row['total_net'], best_match['montant_net_fournisseur'], tolerance)[1],
                    'amount_match': self.calculate_amount_similarity(pdf_row['total_net'], best_match['montant_net_fournisseur'], tolerance)[0],
                    'week_match': False,
                    'commande_match': True
                }
                
                matched_pairs.append(match_data)
                remaining_pdf.remove(pdf_row)
                remaining_beeline.remove(best_match)
                matches_phase2 += 1
        
        st.write(f"✅ Phase 2 terminée: {matches_phase2} matches partiels trouvés")
        
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
            'exact_matches': len([m for m in matched_pairs if m['match_type'] == 'EXACT']),
            'partial_matches': len([m for m in matched_pairs if m['match_type'] == 'PARTIAL'])
        }
        
        return self.matching_stats
    
    def create_consolidated_report(self) -> io.BytesIO:
        """Crée le rapport consolidé Excel"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # Feuille 1: Données consolidées (matches)
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
                    'Semaine_Finissant_Le_PDF': pdf['semaine_finissant_le'],
                    'Semaine_Finissant_Le_Beeline': beeline['semaine_finissant_le'],
                    
                    # Métadonnées de matching
                    'Type_Match': match['match_type'],
                    'Confiance': round(match['confidence'], 3),
                    'Similarite_Montant': round(match['amount_similarity'], 3),
                    'Montants_Coherents': match['amount_match'],
                    'Semaines_Coherentes': match['week_match'],
                    'Ecart_Montant': abs((pdf['total_net'] or 0) - (beeline['montant_net_fournisseur'] or 0)),
                    'Ecart_Pourcentage': round(abs((pdf['total_net'] or 0) - (beeline['montant_net_fournisseur'] or 0)) / max(abs(pdf['total_net'] or 1), abs(beeline['montant_net_fournisseur'] or 1)) * 100, 2)
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
            
            # Feuille 4: Statistiques de matching
            stats_data = [
                ['Métrique', 'Valeur'],
                ['Total lignes PDF', self.matching_stats['total_pdf_rows']],
                ['Total lignes Beeline', self.matching_stats['total_beeline_rows']],
                ['Total matches trouvés', self.matching_stats['total_matched']],
                ['Matches exacts', self.matching_stats['exact_matches']],
                ['Matches partiels', self.matching_stats['partial_matches']],
                ['PDF non-matchés', self.matching_stats['unmatched_pdf']],
                ['Beeline non-matchés', self.matching_stats['unmatched_beeline']],
                ['Taux de match PDF (%)', round(self.matching_stats['match_rate_pdf'], 2)],
                ['Taux de match Beeline (%)', round(self.matching_stats['match_rate_beeline'], 2)]
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
                            'Nb_Factures': 0,
                            'Total_Net_PDF': 0,
                            'Total_Net_Beeline': 0,
                            'Commandes': set(),
                            'Rubriques': set(),
                            'Projets': set()
                        }
                    
                    synthesis = collaborateur_synthesis[collab]
                    synthesis['Nb_Factures'] += 1
                    synthesis['Total_Net_PDF'] += row['Total_Net_PDF'] or 0
                    synthesis['Total_Net_Beeline'] += row['Montant_Net_Fournisseur'] or 0
                    synthesis['Commandes'].add(row['Numero_Commande'])
                    synthesis['Rubriques'].add(row['Code_Rubrique'])
                    synthesis['Projets'].add(row['Projet'])
                
                # Convertir pour export
                synthesis_export = []
                for synthesis in collaborateur_synthesis.values():
                    synthesis_export.append({
                        'Collaborateur': synthesis['Collaborateur'],
                        'Nb_Factures': synthesis['Nb_Factures'],
                        'Total_Net_PDF': synthesis['Total_Net_PDF'],
                        'Total_Net_Beeline': synthesis['Total_Net_Beeline'],
                        'Ecart_Total': synthesis['Total_Net_PDF'] - synthesis['Total_Net_Beeline'],
                        'Nb_Commandes_Uniques': len(synthesis['Commandes']),
                        'Nb_Rubriques_Uniques': len(synthesis['Rubriques']),
                        'Commandes': ', '.join(sorted(synthesis['Commandes'])),
                        'Rubriques': ', '.join(sorted(filter(None, synthesis['Rubriques']))),
                        'Projets': ', '.join(sorted(filter(None, synthesis['Projets'])))
                    })
                
                if synthesis_export:
                    df_synthesis = pd.DataFrame(synthesis_export)
                    df_synthesis.to_excel(writer, sheet_name='Synthese_Collaborateurs', index=False)
        
        output.seek(0)
        return output


def main():
    st.title("🔗 Excel Matcher Beeline")
    st.markdown("### Rapprochement automatique Excel PDF ↔ Excel Beeline")
    
    # Sidebar
    st.sidebar.header("⚙️ Paramètres de matching")
    tolerance = st.sidebar.slider(
        "Tolérance sur les montants (%)", 
        min_value=1, 
        max_value=20, 
        value=5, 
        help="Tolérance acceptée pour considérer que deux montants correspondent"
    ) / 100
    
    st.sidebar.header("📋 Instructions")
    st.sidebar.markdown("""
    1. **Uploadez** vos fichiers Excel PDF (App 1)
    2. **Uploadez** vos fichiers Excel Beeline  
    3. **Lancez** le rapprochement
    4. **Vérifiez** les résultats
    5. **Téléchargez** le rapport consolidé
    """)
    
    st.sidebar.header("🎯 Critères de matching")
    st.sidebar.markdown("""
    - **Exact** : N° commande + Semaine + Montants cohérents
    - **Partiel** : N° commande + Montants similaires
    - **Tolérance** : ±5% par défaut sur les montants
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
    
    # Section 3: Lancement du matching
    if uploaded_excel_pdf and uploaded_excel_beeline:
        st.header("🚀 3. Lancement du rapprochement")
        
        if st.button("🔗 Lancer le rapprochement", type="primary"):
            with st.spinner("Traitement en cours..."):
                
                matcher = ExcelMatcher()
                
                # Phase 1: Chargement des données
                st.subheader("📊 Chargement des données")
                
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
                
                # Phase 2: Rapprochement
                st.subheader("🔗 Rapprochement des données")
                
                matching_stats = matcher.perform_matching(tolerance)
                
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
                    st.metric("✅ Taux match Beeline", f"{matching_stats['match_rate_beeline']:.1f}%")
                
                # Détail des types de matches
                st.subheader("🎯 Types de correspondances")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("🎯 Matches exacts", matching_stats['exact_matches'])
                
                with col2:
                    st.metric("🔍 Matches partiels", matching_stats['partial_matches'])
                
                with col3:
                    st.metric("❌ PDF non-matchés", matching_stats['unmatched_pdf'])
                
                with col4:
                    st.metric("❌ Beeline non-matchés", matching_stats['unmatched_beeline'])
                
                # Aperçu des matches
                if matcher.matched_data:
                    st.subheader("👥 Aperçu des correspondances trouvées")
                    
                    # Tableau des matches
                    matches_display = []
                    for match in matcher.matched_data[:10]:  # Afficher les 10 premiers
                        pdf = match['pdf_data']
                        beeline = match['beeline_data']
                        
                        matches_display.append({
                            'Type': match['match_type'],
                            'Confiance': f"{match['confidence']:.2f}",
                            'N° Facture': pdf['numero_facture'],
                            'N° Commande': pdf['numero_commande'],
                            'Collaborateur': beeline['collaborateur'],
                            'Net PDF (€)': f"{pdf['total_net']:,.2f}" if pdf['total_net'] else "0,00",
                            'Net Beeline (€)': f"{beeline['montant_net_fournisseur']:,.2f}" if beeline['montant_net_fournisseur'] else "0,00",
                            'Écart (€)': f"{abs((pdf['total_net'] or 0) - (beeline['montant_net_fournisseur'] or 0)):,.2f}",
                            'Semaine PDF': pdf['semaine_finissant_le'] or '❌',
                            'Semaine Beeline': beeline['semaine_finissant_le'] or '❌'
                        })
                    
                    df_matches = pd.DataFrame(matches_display)
                    st.dataframe(df_matches, use_container_width=True)
                    
                    if len(matcher.matched_data) > 10:
                        st.info(f"ℹ️ Affichage des 10 premiers matches sur {len(matcher.matched_data)} au total")
                
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
                                'commandes': set()
                            }
                        
                        collaborateurs[collab]['nb_matches'] += 1
                        collaborateurs[collab]['total_net_pdf'] += match['pdf_data']['total_net'] or 0
                        collaborateurs[collab]['total_net_beeline'] += match['beeline_data']['montant_net_fournisseur'] or 0
                        collaborateurs[collab]['commandes'].add(match['pdf_data']['numero_commande'])
                    
                    collab_display = []
                    for name, stats in sorted(collaborateurs.items()):
                        collab_display.append({
                            'Collaborateur': name,
                            'Nb Matches': stats['nb_matches'],
                            'Nb Commandes': len(stats['commandes']),
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
                            for item in matcher.unmatched_pdf[:5]:  # 5 premiers
                                unmatched_pdf_display.append({
                                    'N° Facture': item['numero_facture'],
                                    'N° Commande': item['numero_commande'],
                                    'Total Net (€)': f"{item['total_net']:,.2f}" if item['total_net'] else "0,00",
                                    'Semaine': item['semaine_finissant_le'] or '❌'
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
                            for item in matcher.unmatched_beeline[:5]:  # 5 premiers
                                unmatched_beeline_display.append({
                                    'Collaborateur': item['collaborateur'],
                                    'N° Commande': item['numero_commande'],
                                    'Net Fournisseur (€)': f"{item['montant_net_fournisseur']:,.2f}" if item['montant_net_fournisseur'] else "0,00",
                                    'Semaine': item['semaine_finissant_le'] or '❌'
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
                    - **Donnees_Consolidees** : Toutes les correspondances trouvées
                    - **Non_Matches_PDF** : Données PDF non rapprochées
                    - **Non_Matches_Beeline** : Données Beeline non rapprochées
                    """)
                
                with col2:
                    st.markdown("""
                    **📊 Analyses incluses :**
                    - **Statistiques** : Métriques de performance du matching
                    - **Synthese_Collaborateurs** : Analyse par intérimaire
                    - **Confiance et écarts** : Qualité des rapprochements
                    """)
                
                # Bouton de téléchargement
                filename = f"Rapport_Consolidé_PDF_Beeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                st.download_button(
                    label="📊 Télécharger le rapport consolidé",
                    data=excel_report.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                # Résumé final
                st.success("🎉 Rapprochement terminé avec succès !")
                
                with st.expander("📈 Résumé détaillé", expanded=False):
                    st.markdown(f"""
                    **📊 Données traitées :**
                    - {matching_stats['total_pdf_rows']} lignes Excel PDF
                    - {matching_stats['total_beeline_rows']} lignes Excel Beeline
                    
                    **🔗 Correspondances :**
                    - {matching_stats['exact_matches']} matches exacts
                    - {matching_stats['partial_matches']} matches partiels
                    - {matching_stats['total_matched']} total matches
                    
                    **📈 Performance :**
                    - {matching_stats['match_rate_pdf']:.1f}% des PDF rapprochés
                    - {matching_stats['match_rate_beeline']:.1f}% des Beeline rapprochés
                    
                    **💰 Enrichissement :**
                    - Noms des intérimaires ajoutés aux données PDF
                    - Codes rubriques et détails Beeline intégrés
                    - Cohérence des montants vérifiée
                    """)
    
    else:
        st.info("👆 Commencez par uploader vos fichiers Excel PDF et Beeline")
    
    # Footer
    st.markdown("---")
    st.markdown("**Excel Matcher Beeline** - Version 1.0 | Rapprochement automatique PDF ↔ Beeline")


if __name__ == "__main__":
    main()
