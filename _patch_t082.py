import sys

filepath = 'C:/Cursor/TayfaProject/EquipmentCounter/equipment_counter.py'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Find the function to replace
marker_start = 'def extract_cables_dxf(dxf_path: str) -> list[CableItem]:'
marker_end = """    return sorted(cables.values(), key=lambda c: -c.total_length_m)


# ── Specification table parser"""

start_pos = content.find(marker_start)
end_marker_pos = content.find(marker_end, start_pos)

if start_pos == -1:
    print('ERROR: Could not find start marker')
    sys.exit(1)
if end_marker_pos == -1:
    print('ERROR: Could not find end marker')
    sys.exit(1)

# We want to keep the end marker section comment, just replace up to it
# end_pos points to just after the return statement and blank lines
end_of_func = end_marker_pos + len('    return sorted(cables.values(), key=lambda c: -c.total_length_m)')

print(f'Found function at position {start_pos}, ends at {end_of_func}')
print(f'Replacing {end_of_func - start_pos} characters')

# Write result
with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content[:start_pos])
    f.write(open('C:/Cursor/TayfaProject/EquipmentCounter/_new_func.py', 'r', encoding='utf-8').read())
    f.write(content[end_of_func:])

print('SUCCESS')
