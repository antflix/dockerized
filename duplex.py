# Initialize accumulated_materials as an empty dictionary
accumulated_materials = {}

def calc_duplex(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Single Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Double Barrel MC Connector': 1 * quantity,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'Standard Receptacle': 1 * quantity,
          'Standard Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity
    
    
def calc_decora(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Single Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Double Barrel MC Connector': 1 * quantity,
          'NVent Caddy Mounting Slider Bracket': 0,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'Decora Receptacle': 1 * quantity,
          'Decora Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_gfci(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Single Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Single Barrel MC Connector': 1 * quantity,
          'NVent Caddy Mounting Slider Bracket': 1,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'GFCI Duplex Receptacle': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_cutin  (quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
      
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Single Barrel MC Connector': 1 * quantity,
          'NVent Caddy Mounting Slider Bracket': 0,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'Typical Receptacle': 1 * quantity,
          'Typical Receptacle Cover Plate': 1 * quantity,
          'Cut-In Box': 1 * quantity,
          'Drywall Clamps': 1 * quantity
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_surface(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Box': 2 * quantity,
          '4-Square (Typical) Industrial Cover Plug Plate': 1 * quantity,
          '4-Square Cover': 1 * quantity,
          'Ground Stinger': 2 * quantity,
          '1/4” Toggle Bolts': 6 * quantity,
          '1/2" EMT': 10 * quantity,
          '1/2" EMT Connectors': 2 * quantity,
          '1/2" One Hole Strap': 2 * quantity,
          'Double Barrel MC Connector': 2 * quantity, 
          'Red Heads': 2  * quantity,
          'Wire Nuts': 6 * quantity,
          'Mac-2 Straps': 1 * quantity,
          'KX Straps': 2 * quantity,
          '#12 THHN Black Wire': 15 * quantity,
          '#12 THHN White Wire': 15 * quantity,
          '#12 THHN Green Wire': 15 * quantity,
          'Typical Receptacle': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
          accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_duplex_controlled(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            'Deep 4-Square Bracket Box': 2 * quantity,
            'Single Gang Mud Ring': 1 * quantity,
            'Ground Stinger': 2 * quantity,
            'Tek Screws': 10 * quantity,
            'Mac-2 Straps': 3 * quantity,
            'Wire Nuts': 5 * quantity,
            'Red Heads': 4 * quantity,
            'Double Barrel MC Connector': 1 * quantity,
            'Single Barrel MC Connector': 2 * quantity,
            'NVent Caddy Mounting Slider Bracket': 2 * quantity,
            '12/2 LV MC': 20 * quantity,
            '12/3 LV MC': 10 * quantity,
            'Half-Duplex Controlled (Typical) Outlet': 1 * quantity,
            'Typical Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

    

# ----------------------------------------------------------------------

def calc_quad(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Single Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Double Barrel MC Connector': 1 * quantity,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'Standard Receptacle': 2 * quantity,
          'Standard Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity
    
    
def calc_quad_decora(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Single Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Double Barrel MC Connector': 1 * quantity,
          'NVent Caddy Mounting Slider Bracket': 0,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'Decora Receptacle': 2 * quantity,
          'Two Gang Decora Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_quad_gfci(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Single Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Double Barrel MC Connector': 1 * quantity,
          'NVent Caddy Mounting Slider Bracket': 1,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'GFCI Duplex Receptacle': 2 * quantity,
          'Two Gang Decora Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_quad_cutin  (quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          'Ground Stinger': 1 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 1 * quantity,
          'Red Heads': 2 * quantity,
          'Double Barrel MC Connector': 1 * quantity,
          'Single Barrel MC Connector': 2 * quantity,
          '12/2 LV MC': 30 * quantity,
          'KX Straps': 2 * quantity,
          'Typical Receptacle': 2 * quantity,
          'Typical Receptacle Cover Plate': 1 * quantity,
          'Cut-In Box': 2 * quantity,
          'Drywall Clamps': 1 * quantity
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_quad_surface(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
          '4-Square Box': 2 * quantity,
          '2 Gang (Typical) Industrial Cover Plug Plate': 1 * quantity,
          '4-Square Cover': 1 * quantity,
          'Ground Stinger': 2 * quantity,
          '1/4” Toggle Bolts': 6 * quantity,
          '1/2" EMT': 10 * quantity,
          '1/2" EMT Connectors': 2 * quantity,
          '1/2" One Hole Strap': 2 * quantity,
          'Double Barrel MC Connector': 2 * quantity, 
          'Red Heads': 2  * quantity,
          'Wire Nuts': 6 * quantity,
          'KX Straps': 2 * quantity,
          '#12 THHN Black Wire': 15 * quantity,
          '#12 THHN White Wire': 15 * quantity,
          '#12 THHN Green Wire': 15 * quantity,
          'Typical Receptacle': 2 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
          accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_quad_controlled(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            'Deep 4-Square Bracket Box': 2 * quantity,
            'Two Gang Mud Ring': 1 * quantity,
            'Ground Stinger': 2 * quantity,
            'Tek Screws': 10 * quantity,
            'Mac-2 Straps': 6 * quantity,
            'Wire Nuts': 5 * quantity,
            'Red Heads': 4 * quantity,
            'Double Barrel MC Connector': 1 * quantity,
            'Single Barrel MC Connector': 2 * quantity,
            'NVent Caddy Mounting Slider Bracket': 3 * quantity,
            '12/2 LV MC': 20 * quantity,
            '12/3 LV MC': 10 * quantity,
            'Full-Duplex Controlled (Typical) Outlet': 1 * quantity,
            'Typical Receptacle Cover Plate': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

    # -------------------------------------------------------------------------

def calc_ff3(quantity):
  global accumulated_materials
  quantity = int(quantity)  # Convert the quantity to an integer
  if quantity != 0:
      materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Two Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 8 * quantity,
          'Red Heads': 1 * quantity,
          'Single Barrel MC Connector': 1 * quantity,
          '12/3 LV MC': 30 * quantity,
          'Two Gang Stainless Steel Blank Plate': 1 * quantity,
          '1/2” 90 Degree Flex Connector': 1 * quantity,
      }
      # Update accumulated_materials with the quantities from materials
      for item, quantity in materials.items():
          accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_ff4(quantity):
  global accumulated_materials
  quantity = int(quantity)  # Convert the quantity to an integer
  if quantity != 0:
      materials = {
          '4-Square Bracket Box': 1 * quantity,
          'Two Gang Mud Ring': 1 * quantity,
          'Ground Stinger': 1 * quantity,
          'Tek Screws': 10 * quantity,
          'Mac-2 Straps': 4 * quantity,
          'Wire Nuts': 8 * quantity,
          'Red Heads': 4 * quantity,
          'Double Barrel MC Connector': 2 * quantity,
          '12/3 LV MC': 30 * quantity,
          'Two Gang Stainless Steel Blank Plate': 1 * quantity,
          '1/2” 90 Degree Flex Connector': 1 * quantity,
          'NVent Caddy Mounting Slider Bracket': 1 * quantity,
      }
      # Update accumulated_materials with the quantities from materials
      for item, quantity in materials.items():
          accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_rough_data(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '4-Square Bracket Box': 1 * quantity,
            'Single Gang Mud Ring': 1 * quantity,
            'Tek Screws': 4 * quantity,
            'Jet Line': 10 * quantity,
            '3/4” Snap-In Bushings': 1 * quantity,
            '1" Snap-In Bushings': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_cutin_data(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            'MP1s': 1 * quantity,
            'Jet Line': 10 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity


def calc_lv_switch(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '4-Square Bracket Box': 1 * quantity,
            'Single Gang Mud Ring': 1 * quantity,
            'Jet Line': 10 * quantity,
            '3/4” Snap-In Bushings': 1 * quantity,
            'Deep 4-Square Bracket Box': 1 * quantity,
            'Double Barrel MC Connector': 2 * quantity,
            'Tek Screws': 6 * quantity,
            'KX Straps': 4 * quantity,
            '4-Square Cover': 1 * quantity,
            'Ground Stinger': 1 * quantity,
            'Wire Nuts': 1 * quantity,
            '12/2 HV MC': 15 * quantity,
            'Ceiling Wires': 1 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_hv_switch(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '12/3 HV MC': 10 * quantity,
            'Mac-2 Straps': 3 * quantity,
            'Ground Stinger': 2 * quantity,
            'Deep 4-Square Bracket Box': 1 * quantity,
            'Double Barrel MC Connector': 2 * quantity,
            'Single Barrel MC Connector': 1 * quantity,
            'Red Heads': 1 * quantity,
            'Tek Screws': 6 * quantity,
            'Wire Nuts': 3 * quantity,
          
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity


def calc_hv_dimming(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '12/3 HV MC': 10 * quantity,
            'Mac-2 Straps': 3 * quantity,
            'Ground Stinger': 1 * quantity,
            'Single Barrel MC Connector': 1 * quantity,
            'Deep 4-Square Bracket Box': 1 * quantity,
            'Double Barrel MC Connector': 2 * quantity,
            'Red Heads': 1 * quantity,
            'Tek Screws': 6 * quantity,
            'Wire Nuts': 3 * quantity,
            '18/2 LV Dimmer Cable': 15 * quantity,
            'Blue/Orange Wire Nuts': 2 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity


def calc_wh_120(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '3/4" MC Connector': 4 * quantity,
            'Deep 4-Square Box': 1 * quantity,
            '4-Square Industrial Switch Cover Plate': 1 * quantity,
            '1/2" One Hole Strap': 3 * quantity,
            'Big Blue Wire Nuts': 4 * quantity,
            'Two-Pole Motor Rated 40A Switch': 1 * quantity,
            '8/3 LV MC': 30 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

def calc_wh_277(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '3/4" MC Connector': 4 * quantity,
            'Deep 4-Square Box': 1 * quantity,
            '4-Square Industrial Switch Cover Plate': 1 * quantity,
            '1/2" One Hole Strap': 3 * quantity,
            'Big Blue Wire Nuts': 4 * quantity,
            'Single-Pole Motor Rated 30A Switch': 1 * quantity,
            '10/2 HV MC': 30 * quantity,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity



            
            
            
            
            
            
            
def calc_lights(quantity):
    global accumulated_materials
    quantity = int(quantity)  # Convert the quantity to an integer
    if quantity != 0:
        materials = {
            '4-Square Bracket Box': 1 * quantity,
            '4" Octagon Box': 0,
            'Ceiling Fixture Box': 1 * quantity,
            'Ground Stinger': 1 * quantity,
            'Tek Screws': 10 * quantity,
            'Mac-2 Straps': 4 * quantity,
            'Wire Nuts': 1 * quantity,
            'Red Heads': 2 * quantity,
            'Double Barrel MC Connector': 1 * quantity,
            '12/2 LV MC': 15 * quantity,
            'Light Fixture': 1 * quantity,
            'Light Fixture Cover Plate': 0,
        }
        # Update accumulated_materials with the quantities from materials
        for item, quantity in materials.items():
            accumulated_materials[item] = accumulated_materials.get(item, 0) + quantity

   


def reset_variables():
    global accumulated_materials
    keys_to_remove = list(accumulated_materials.keys())  # Create a copy of keys
    for item in keys_to_remove:
        accumulated_materials[item] = 0
    keys_to_remove = []
    for item, quantity in accumulated_materials.items():
        if quantity == 0:
            keys_to_remove.append(item)
    for item in keys_to_remove:
        del accumulated_materials[item]