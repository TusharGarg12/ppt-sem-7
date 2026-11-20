with open('scratch/generate_build_deck_with_references.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Replace border loops that went up to $cell.Borders.Count with 1..4
# Replace 0x881414 with 0x1C1CB7 for crimson fill
# Ensure $cell.Shape.Fill.Transparency = 0.0 is present

# Slide 4 snippet to replace/ensure
old_s4_block = '''        for ($b = 1; $b -le $cell.Borders.Count; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        
        if ($r -eq 0) {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0x881414'''

new_s4_block = '''        for ($b = 1; $b -le 4; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        
        $cell.Shape.Fill.Transparency = 0.0
        $cell.Shape.Fill.Solid()
        
        if ($r -eq 0) {
            $cell.Shape.Fill.ForeColor.RGB = 0x1C1CB7'''

if old_s4_block in text:
    text = text.replace(old_s4_block, new_s4_block)
    print("Replaced Slide 4 border & fill block!")
else:
    print("WARNING: old_s4_block not found!")

# Slide J (17) snippet to replace/ensure
old_sj_block = '''        for ($b = 1; $b -le $cell.Borders.Count; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        
        if ($r -eq 0) {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0x881414'''

new_sj_block = '''        for ($b = 1; $b -le 4; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        
        $cell.Shape.Fill.Transparency = 0.0
        $cell.Shape.Fill.Solid()
        
        if ($r -eq 0) {
            $cell.Shape.Fill.ForeColor.RGB = 0x1C1CB7'''

if old_sj_block in text:
    text = text.replace(old_sj_block, new_sj_block)
    print("Replaced Slide J border & fill block!")
else:
    print("WARNING: old_sj_block not found!")

# Also ensure title for Slide 4 is "Literature Survey: Dining Demand & Food Systems"
text = text.replace('"Literature Survey: Food Systems & Demand Forecasting"', '"Literature Survey: Dining Demand & Food Systems"')

with open('scratch/generate_build_deck_with_references.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("generate_build_deck_with_references.py updated successfully!")
