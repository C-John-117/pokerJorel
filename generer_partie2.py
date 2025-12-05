import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Cr%C3%A9ation du workbook
wb = Workbook()
ws = wb.active
ws.title = "Partie 2"

# Style pour les titres
title_font = Font(bold=True, size=14, color="FFFFFF")
title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True)

# PARTIE 2 - TITRE
ws['B1'] = 'PARTIE 2'
ws['B1'].font = title_font
ws['B1'].fill = title_fill
ws.merge_cells('B1:F1')

# 1) Calcul des coordonnées des sommets des polyomino
ws['B2'] = '1) Calcul des coordonnées des sommets des polyomino'
ws['B2'].font = header_font
ws.merge_cells('B2:E2')

# P2 - 7 cubes poly 0
row = 4
ws[f'B{row}'] = '(P2) Coordonnée : 7 cubes poly 0'
ws[f'B{row}'].font = header_font
ws.merge_cells(f'B{row}:D{row}')

row += 2
ws[f'B{row}'] = 'Point'
ws[f'C{row}'] = 'x'
ws[f'D{row}'] = 'y'
for cell in [ws[f'B{row}'], ws[f'C{row}'], ws[f'D{row}']]:
    cell.font = header_font

# Coordonnées P2
p2_coords = [
    ('A', 2, 0),
    ('B', 3, 0),
    ('C', 3, -1),
    ('', 5, -1),
    ('D', 5, 0),
    ('E', 6, 0),
    ('F', 6, -3),
    ('G', 5, -3),
    ('H', 5, -2),
    ('I', 2, -2),
    ('A', 2, 0)
]

row += 1
for point, x, y in p2_coords:
    ws[f'B{row}'] = point
    ws[f'C{row}'] = x
    ws[f'D{row}'] = y
    row += 1

# P1 - 7 cubes poly 2
row += 2
ws[f'B{row}'] = '(P1) Coordonnée : 7 cubes poly 2'
ws[f'B{row}'].font = header_font
ws.merge_cells(f'B{row}:D{row}')

row += 2
ws[f'B{row}'] = 'Point'
ws[f'C{row}'] = 'X'
ws[f'D{row}'] = 'Y'
for cell in [ws[f'B{row}'], ws[f'C{row}'], ws[f'D{row}']]:
    cell.font = header_font

# Coordonnées P1
p1_coords = [
    ('K', 2, 0),
    ('L', 3, 0),
    ('M', 3, -1),
    ('N', 5, -1),
    ('O', 5, -2),
    ('P', 6, -2),
    ('Q', 6, -3),
    ('R', 3, -3),
    ('S', 3, -2),
    ('T', 2, -2),
    ('K', 2, 0)
]

row += 1
for point, x, y in p1_coords:
    ws[f'B{row}'] = point
    ws[f'C{row}'] = x
    ws[f'D{row}'] = y
    row += 1

# 2) Matrices de transformation pour les rotations
row += 2
ws[f'B{row}'] = '2) Calcul des matrices de transformation pour les rotations'
ws[f'B{row}'].font = header_font
ws.merge_cells(f'B{row}:F{row}')

# Fonction pour ajouter une matrice
def add_matrix(ws, row, label, matrix):
    ws[f'B{row}'] = label
    ws[f'B{row}'].font = Font(color="C00000")
    for i in range(3):
        for j in range(3):
            ws.cell(row=row+i, column=3+j, value=matrix[i][j])
    return row + 4

# Matrices de rotation
row += 2
rotation_90 = [[0, -1, 0], [1, 0, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Rotation 90', rotation_90)

rotation_180 = [[-1, 0, 0], [0, -1, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Rotation 180', rotation_180)

rotation_270 = [[0, 1, 0], [-1, 0, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Rotation 270', rotation_270)

rotation_m180 = [[-1, 0, 0], [0, -1, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Rotation -180', rotation_m180)

rotation_m270 = [[0, -1, 0], [1, 0, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Rotation -270', rotation_m270)

rotation_m90 = [[0, 1, 0], [-1, 0, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Rotation -90', rotation_m90)

# Homothéties
ws[f'B{row}'] = '2) Matrice de transformation pour une homothétie de facteur 2 et de facteur -2'
ws[f'B{row}'].font = header_font
ws.merge_cells(f'B{row}:F{row}')

row += 2
homothetie_2 = [[2, 0, 0], [0, 2, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Homothétie de facteur 2', homothetie_2)

homothetie_m2 = [[-2, 0, 0], [0, -2, 0], [0, 0, 1]]
row = add_matrix(ws, row, 'Homothétie de facteur -2', homothetie_m2)

# Translations
ws[f'B{row}'] = '2) Calcul des matrices de translation pour de k vers la gauche, k vers la droite, k vers haut et K vers le bas.'
ws[f'B{row}'].font = header_font
ws.merge_cells(f'B{row}:F{row}')

row += 2
ws[f'B{row}'] = 'Translation de k vers la gauche'
ws.cell(row=row, column=3, value=1)
ws.cell(row=row, column=4, value=0)
ws.cell(row=row, column=5, value='-k')
ws.cell(row=row+1, column=3, value=0)
ws.cell(row=row+1, column=4, value=1)
ws.cell(row=row+1, column=5, value=0)
ws.cell(row=row+2, column=3, value=0)
ws.cell(row=row+2, column=4, value=1)
ws.cell(row=row+2, column=5, value=1)
row += 4

ws[f'B{row}'] = 'Translation de K vers la droite'
ws.cell(row=row, column=3, value=1)
ws.cell(row=row, column=4, value=0)
ws.cell(row=row, column=5, value='k')
ws.cell(row=row+1, column=3, value=0)
ws.cell(row=row+1, column=4, value=1)
ws.cell(row=row+1, column=5, value=0)
ws.cell(row=row+2, column=3, value=0)
ws.cell(row=row+2, column=4, value=0)
ws.cell(row=row+2, column=5, value=1)
row += 4

ws[f'B{row}'] = 'Translation de K vers le haut'
ws.cell(row=row, column=3, value=1)
ws.cell(row=row, column=4, value=0)
ws.cell(row=row, column=5, value=0)
ws.cell(row=row+1, column=3, value=0)
ws.cell(row=row+1, column=4, value=1)
ws.cell(row=row+1, column=5, value='k')
ws.cell(row=row+2, column=3, value=0)
ws.cell(row=row+2, column=4, value=0)
ws.cell(row=row+2, column=5, value=1)
row += 4

ws[f'B{row}'] = 'Translation de K vers le bas'
ws.cell(row=row, column=3, value=1)
ws.cell(row=row, column=4, value=0)
ws.cell(row=row, column=5, value=0)
ws.cell(row=row+1, column=3, value=0)
ws.cell(row=row+1, column=4, value=1)
ws.cell(row=row+1, column=5, value='-k')
ws.cell(row=row+2, column=3, value=0)
ws.cell(row=row+2, column=4, value=0)
ws.cell(row=row+2, column=5, value=1)

# Sauvegarder le fichier
wb.save('partie2_polyominos.xlsx')
print("Fichier Excel g%C3%A9n%C3%A9r%C3%A9 : partie2_polyominos.xlsx")