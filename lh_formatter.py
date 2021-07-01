import json
import xlsxwriter

# This script is to convert a Log Horizon TRPG monster from .json to .txt in a format friendly for Discord.

def bold(text):
    return ("**" + text + "**")

def ital(text):
    return ("*" + text + "*")

# Discord only accepts underline formatting after bold/italic, so call last.
def undl(text):
    return ("__" + text + "__")


in_file = open("input.json", "r")
data = json.load(in_file)
in_file.close()

out_file = open("output.txt", "w")

for x in data['Monsters'].keys():
    
    output =  "=" * 50 + "\n" + ">>> "
    output += undl(bold(data['Monsters'][x]['Name'])) + "\n"

    output += bold("Tags") + ": "
    for tag in data['Monsters'][x]['Tags']:
        output += "[" + tag + "]"
    output += "\n"

    output += bold("Rank") + ": " + data['Monsters'][x]['Rank'] + " " * 3
    output += bold("HP") + ": " + data['Monsters'][x]['HP'] + "\n"

    output += bold("STR") + ": " + data['Monsters'][x]['STR'] + "\t"
    output += bold("DEX") + ": " + data['Monsters'][x]['DEX'] + "\t"
    output += bold("POW") + ": " + data['Monsters'][x]['POW'] + "\t"
    output += bold("INT") + ": " + data['Monsters'][x]['INT'] + "\n"

    output += bold("Evasion") + ": " + data['Monsters'][x]['Evasion'] + " " * 22
    output += bold("Resistance") + ": " + data['Monsters'][x]['Resistance'] + "\n"

    output += bold("Physical Defense") + ": " + data['Monsters'][x]['Physical Defense'] + "\t"
    output += bold("Magic Defense") + ": " + data['Monsters'][x]['Magic Defense'] + "\n"

    output += bold("Movement Speed") + ": " + data['Monsters'][x]['Movement Speed'] + " " * 3
    output += bold("Initiative") + ": " + data['Monsters'][x]['Initiative'] + "\n"

    output += bold("Identification Difficulty") + ": " + data['Monsters'][x]['Identification Difficulty'] + "\t"
    output += bold("Hate Multiplier") + ": " + data['Monsters'][x]['Hate Multiplier'] + "\n"

    for ability in data['Monsters'][x]['Abilities']:
        output += "\n" + undl(bold(ability['Name'])) + "\n"

        if ability['Tags']:
            output += bold("Tags") + ": "
            for tag in ability['Tags']:
                output += "[" + tag + "]"
            output += "\n"

        output += bold("Timing") + ": " + ability['Timing'] + "\t"
        output += bold("Limit") + ": " + ability['Limit'] + "\n"

        if ability['Range']:
            output += bold("Range") + ": " + ability['Range'] + "\t"

        if ability['Check']:
            output += bold("Check") + ": " + ability['Check']  + "\n"
        output += bold("Effect") + ": " + ability['Effect'] + "\n"

        if ability['Extra']:
            output +=  ability['Extra'] + "\n"

    output += "\n" + undl(bold("Drops")) + ": \n"
    for drops in data['Monsters'][x]['Drops']:
            output += drops + ": " + data['Monsters'][x]['Drops'][drops] + "\n"

    output += "-" * 50 + "\n"

    # Now for the identifiable information

    output += ">>> " 
    output += undl(bold(data['Monsters'][x]['Name'] + " (Identified)")) + "\n"
    output += bold("Rank") + ": " + data['Monsters'][x]['Rank'] + " " * 3
    output += bold("Tags") + ": "
    for tag in data['Monsters'][x]['Tags']:
        output += "[" + tag + "]"
    output += "\n"

    output += bold("Defense") + ": "
    if data['Monsters'][x]['Physical Defense'] and data['Monsters'][x]['Magic Defense']:
        if int(data['Monsters'][x]['Physical Defense']) > int(data['Monsters'][x]['Magic Defense']):
            output += "Physical Defense is superior.\n"
        elif int(data['Monsters'][x]['Physical Defense']) < int(data['Monsters'][x]['Magic Defense']):
            output += "Magic Defense is superior.\n"
        else:
            output += "Defenses are equal.\n"
    else:
        output += "One or both defense values were null.\n"

    output += bold("Hate Multiplier") + ": " + data['Monsters'][x]['Hate Multiplier'] + "\n"

    for ability in data['Monsters'][x]['Abilities']:
        output += "\n" + undl(bold(ability['Name'])) + "\n"

        if ability['Tags']:
            output += bold("Tags") + ": "
            for tag in ability['Tags']:
                output += "[" + tag + "]"
            output += "\n"

        output += bold("Timing") + ": " + ability['Timing'] + "\t"
        output += bold("Limit") + ": " + ability['Limit'] + "\n"

        if ability['Range']:
            output += bold("Range") + ": " + ability['Range'] + "\t"

        if ability['Check']:
            output += bold("Check") + ": " + ability['Check']  + "\n"
        output += bold("Effect") + ": " + ability['Effect'] + "\n"

        if ability['Extra']:
            output += ability['Extra'] + "\n"

    
    # Unidentified portion

    output += "\n" + "-" * 50 + "\n" + ">>> " 
    output += undl(bold(data['Monsters'][x]['Name'] + " (Unidentified)")) + "\n"
    output += bold("Rank") + ": " + data['Monsters'][x]['Rank'] + " " * 3
    output += bold("Tags") + ": "
    for tag in data['Monsters'][x]['Tags']:
        output += "[" + tag + "]"
    output += "\n"

    out_file.write(output)

# Props
for x in data['Props'].keys():
    
    output =  "=" * 50 + "\n" + ">>> "
    output += undl(bold(data['Props'][x]['Name'])) + "\n"
    output += bold("Rank") + ": " + data['Props'][x]['Rank'] + "\n"
    for tag in data['Props'][x]['Tags']:
        output += "[" + tag + "]"
    output += "\n"
    output += bold("Detect") + ": " + data['Props'][x]['Detect'] + " " * 3
    output += bold("Analyse") + ": " + data['Props'][x]['Analyse'] + " " * 3
    output += bold("Disable") + ": " + data['Props'][x]['Disable'] + "\n"
    output += bold("Effect") + ": " + data['Props'][x]['Effect'] + "\n"
    output +=  "=" * 50 + "\n"
    out_file.write(output)

out_file.close()

# Excel workbork

workbook = xlsxwriter.Workbook("Monsters.xlsx")
worksheet = workbook.add_worksheet()
worksheet.set_column_pixels(0, 50, 120)
# somehow, the set row is per individual row?

x = 1000
while x < 1000:
    worksheet.set_row_pixels(x, 21)
    x += 1

# Create the black fill, white text format for headers
header = workbook.add_format()
header.set_bg_color('black')
header.set_font_name('Arial')
header.set_font_size(10)
header.set_font_color('white')
header.set_border(1)
header.set_bold()
header.set_align('center')
header.set_align('vcenter')

# Create body text for statistics

stats = workbook.add_format()
stats.set_font_name('Arial')
stats.set_font_size(10)
stats.set_text_wrap()
stats.set_border(1)
stats.set_align('center')
stats.set_align('vcenter')

# Create body text for abilities

body = workbook.add_format()
body.set_font_name('Arial')
body.set_font_size(10)
body.set_text_wrap()
body.set_align('vcenter')
body.set_border(1)

# Create drop table text

drop = workbook.add_format()
drop.set_font_name('Arial')
drop.set_font_size(10)
drop.set_text_wrap()
drop.set_border(1)
drop.set_align('center')
drop.set_align('vcenter')

# make a crawler, we should start from B3

og_row = 2
og_col = 1
height = 0
i = 0

for x in data['Monsters'].keys():

    row = og_row
    col = og_col
    
    row += height

    worksheet.merge_range(row, col, row, col + 5, "")
    worksheet.write(row, col, data['Monsters'][x]['Name'], header)
    row += 1
    height += 1

    worksheet.write(row, col, "Tags", header)
    worksheet.merge_range(row, col + 1, row, col + 5, "")
    tag_str = ""
    for tag in data['Monsters'][x]['Tags']:
        tag_str += "[" + tag + "]" 
    worksheet.write(row, col + 1, tag_str, stats) 
    row += 1
    height += 1

    worksheet.write(row, col, "Rank", header)
    worksheet.write(row, col + 1, data['Monsters'][x]['Rank'], stats)

    worksheet.write(row, col + 2, "STR", header)
    worksheet.write(row, col + 3, data['Monsters'][x]['STR'], stats)

    worksheet.write(row, col + 4, "Evasion", header)
    worksheet.write(row, col + 5, data['Monsters'][x]['Evasion'], stats)

    row += 1
    height += 1

    worksheet.write(row, col, "HP", header)
    worksheet.write(row, col + 1, data['Monsters'][x]['HP'], stats)

    worksheet.write(row, col + 2, "DEX", header)
    worksheet.write(row, col + 3, data['Monsters'][x]['DEX'], stats)

    worksheet.write(row, col + 4, "Resistance", header)
    worksheet.write(row, col + 5, data['Monsters'][x]['Resistance'], stats)

    row += 1
    height += 1

    worksheet.write(row, col, "Initiative", header)
    worksheet.write(row, col + 1, data['Monsters'][x]['Initiative'], stats)

    worksheet.write(row, col + 2, "POW", header)
    worksheet.write(row, col + 3, data['Monsters'][x]['POW'], stats)

    worksheet.write(row, col + 4, "Phys. Def", header)
    worksheet.write(row, col + 5, data['Monsters'][x]['Physical Defense'], stats)

    row += 1
    height += 1

    worksheet.write(row, col, "Speed", header)
    worksheet.write(row, col + 1, data['Monsters'][x]['Movement Speed'], stats)

    worksheet.write(row, col + 2, "INT", header)
    worksheet.write(row, col + 3, data['Monsters'][x]['INT'], stats)

    worksheet.write(row, col + 4, "Mag. Def", header)
    worksheet.write(row, col + 5, data['Monsters'][x]['Magic Defense'], stats)

    row += 1
    height += 1

    worksheet.merge_range(row, col, row, col + 1, "", stats)
    worksheet.write(row, col, "Identification Difficulty", header)
    worksheet.write(row, col + 2, data['Monsters'][x]['Identification Difficulty'], stats)

    worksheet.merge_range(row, col + 3, row, col + 4, "", stats)
    worksheet.write(row, col + 3, "Hate Multiplier", header)
    worksheet.write(row, col + 5, data['Monsters'][x]['Hate Multiplier'], stats)

    row += 1
    height += 1

    

    for ability in data['Monsters'][x]['Abilities']:
        worksheet.merge_range(row, col, row, col + 5, "", header)
        worksheet.write(row, col, ability['Name'], header)

        row += 1
        height += 1

        if ability['Tags']:
            tag_str = ""
            for tag in ability['Tags']:
                tag_str += "[" + tag + "]"
            tag += "\n"
            worksheet.merge_range(row, col, row, col + 5, "", stats)
            worksheet.write(row, col, tag_str, stats)

            row += 1
            height += 1

        worksheet.merge_range(row, col, row, col + 1, "", header)
        worksheet.write(row, col, "Timing", header)
        worksheet.write(row, col + 2, ability['Timing'], stats)

        worksheet.merge_range(row, col + 3, row, col + 4, "", header)
        worksheet.write(row, col + 3, "Limit", header)
        worksheet.write(row, col + 5, ability['Limit'], stats)

        row += 1
        height += 1

        if ability['Range']:
            worksheet.merge_range(row, col, row, col + 1, "", header)
            worksheet.write(row, col, "Range", header)
            worksheet.write(row, col + 2, ability['Range'], stats)

        if ability['Check']:
            worksheet.merge_range(row, col + 3, row, col + 4, "", header)
            worksheet.write(row, col + 3, "Check", header)
            worksheet.write(row, col + 5, ability['Check'], stats)
            worksheet.set_row_pixels(row, ((len(ability['Check']) // 13)) * 20)

        if ability['Range'] or ability['Check']:
            row += 1
            height += 1

        ability_str = "Effect: " + ability['Effect']
        
        effect_row_height = 0

        if ability['Extra']:
            ability_str += "\n" + ability['Extra']
            effect_row_height = (len(("Extra: " + ability['Extra'])) // 100)

        effect_row_height += (len("Effect: " + ability['Effect']) // 100) + 1
        

        worksheet.merge_range(row, col, row, col + 5, "", body)
        worksheet.write(row, col, ability_str, body)
        worksheet.set_row_pixels(row, (ability_str.count("\n") + effect_row_height) * 20)

        row += 1
        height += 1

    worksheet.merge_range(row, col, row, col + 5, "", header)
    worksheet.write(row, col, "Drop Items", header)

    row += 1
    height += 1

    for drops_data in data['Monsters'][x]['Drops']:
        worksheet.write(row, col, drops_data, drop)
        worksheet.merge_range(row, col + 1, row, col + 5, "", drop)
        worksheet.write(row, col + 1, data['Monsters'][x]['Drops'][drops_data], drop)
        row += 1
        height += 1
        

    height += 1

for x in data['Props'].keys():
    row = og_row
    col = og_col
    
    row += height

    worksheet.merge_range(row, col, row, col + 5, "", header)
    worksheet.write(row, col, data['Props'][x]['Name'], header)

    row += 1
    height += 1

    worksheet.write(row, col, "Rank", header)
    worksheet.write(row, col + 1, data['Props'][x]['Rank'], stats)

    worksheet.merge_range(row, col + 2 , row, col + 5, "", stats)
    tag_str = ""
    for tag in data['Props'][x]['Tags']:
        tag_str += "[" + tag + "]" 
    worksheet.write(row, col + 2, tag_str, stats)

    row += 1
    height += 1

    worksheet.write(row, col, "Detect", header)
    worksheet.write(row, col + 1, data['Props'][x]['Detect'], stats)
    
    worksheet.write(row, col + 2, "Analyse", header)
    worksheet.write(row, col + 3, data['Props'][x]['Analyse'], stats)
    
    worksheet.write(row, col + 4, "Disable", header)
    worksheet.write(row, col + 5, data['Props'][x]['Disable'], stats)

    row += 1
    height += 1

    worksheet.merge_range(row, col , row, col + 5, "", body)
    worksheet.write(row, col, data['Props'][x]['Effect'], body)
    worksheet.set_row_pixels(row, ((len(data['Props'][x]['Effect']) // 100) + 1)* 20)

    row += 1
    height += 1


workbook.close()
