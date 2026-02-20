import pandas as pd

def main():
    # 1. Načtení dat (veškerá data načteme jako text, aby nám nezmizely úvodní nuly)
    # Změň název souboru na tvůj aktuální, pokud se jmenuje jinak
    file_path = "pick___.XLSX - Sheet1.csv"
    df = pd.read_csv(file_path, dtype=str)
    
    # Odebereme řádky, kde chybí číslo zakázky (Delivery)
    df = df.dropna(subset=['Delivery']).copy()
    
    # 2. Převedeme množství na číselný typ pro pozdější součty
    # Pro počet kusů obvykle používáme 'Act.qty (dest)' nebo 'Source actual qty.'
    df['Qty'] = pd.to_numeric(df['Act.qty (dest)'], errors='coerce').fillna(0)

    # 3. Agregace dat za každou zakázku (Delivery)
    grouped = df.groupby('Delivery').agg(
        num_materials=('Material', 'nunique'),                   # počet unikátních materiálů
        certs=('Certificate Number', lambda x: x.dropna().unique().tolist()), # seznam certifikátů v zakázce
        total_qty=('Qty', 'sum'),                                # celkový počet kusů na zakázku
        num_positions=('Source Storage Bin', 'nunique')          # počet unikátních pickovacích pozic
    )

    # Funkce pro ověření podmínky certifikátu
    def is_valid_cert(certs):
        # Odstranění případných 'nan' a prázdných hodnot
        valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
        
        # Nesmí být prázdné
        if len(valid_certs) == 0:
            return False
            
        # Žádný z certifikátů nesmí začínat na '460'
        for c in valid_certs:
            if c.startswith('460'):
                return False
                
        # Pokud splňuje, prošlo to filtrem
        return True

    # 4. Aplikace filtrů
    # Chceme pouze 1 materiál A zároveň musí splňovat naši certifikátovou podmínku
    filtered_orders = grouped[
        (grouped['num_materials'] == 1) & 
        (grouped['certs'].apply(is_valid_cert))
    ]

    # 5. Výpočet průměrů
    # Pokud bys chtěl jako "počet pozic" počítat počet řádků v excelu na danou zakázku, 
    # dal by se místo 'nunique' (viz výše) použít 'count'
    avg_qty = filtered_orders['total_qty'].mean()
    avg_pos = filtered_orders['num_positions'].mean()
    
    # Volitelně spočítáme i celkový počet vyfiltrovaných zakázek
    total_filtered_orders = len(filtered_orders)

    # 6. Vypsání výsledků
    print("-" * 50)
    print(f"Počet nalezených zakázek (1 materiál, na paletu): {total_filtered_orders}")
    if total_filtered_orders > 0:
        print(f"Průměrný počet kusů na zakázku: {avg_qty:.2f}")
        print(f"Průměrný počet pickovacích pozic na zakázku: {avg_pos:.2f}")
    print("-" * 50)

if __name__ == "__main__":
    main()
