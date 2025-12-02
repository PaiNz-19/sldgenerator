import io
import math
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.patches as patches

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ============================================================
# 1. UTILS & CONFIGURATION
# ============================================================

def get_tx_chain(voltage, scheme):
    if voltage == "400V": return []
    if voltage == "11kV": return [{"ratio": "11/0.4 kV", "bus": "0.4kV"}]
    if voltage == "33kV": return [{"ratio": "33/0.4 kV", "bus": "0.4kV"}]
    if voltage == "132kV":
        if scheme == "132/0.4 kV": return [{"ratio": "132/0.4 kV", "bus": "0.4kV"}]
        elif scheme == "132/11/0.4 kV": return [{"ratio": "132/11 kV", "bus": "11kV"}, {"ratio": "11/0.4 kV", "bus": "0.4kV"}]
        else: return [{"ratio": "132/33 kV", "bus": "33kV"}, {"ratio": "33/11 kV", "bus": "11kV"}, {"ratio": "11/0.4 kV", "bus": "0.4kV"}]
    return []

def get_lv_gen_inputs(key_prefix, include_emsb=False):
    if include_emsb:
        c1, c2, c3 = st.columns(3)
    else:
        c1, c2 = st.columns(2)
        
    has_solar = c1.checkbox("Solar PV", key=f"{key_prefix}_sol")
    has_bess = c2.checkbox("BESS", key=f"{key_prefix}_bess")
    
    has_emsb = False
    if include_emsb:
        has_emsb = c3.checkbox("EMSB", key=f"{key_prefix}_emsb")
    
    gens = []
    if has_solar:
        st.markdown("**Solar PV Specs**")
        ca, cb = st.columns(2)
        kwac = ca.number_input("kWac", 0, 99999, 100, key=f"{key_prefix}_skwa")
        kwp = cb.number_input("kWp", 0, 99999, 120, key=f"{key_prefix}_skwp")
        gens.append({"type": "Solar", "kWac": kwac, "cap_val": kwp})
        
    if has_bess:
        st.markdown("**BESS Specs**")
        ca, cb = st.columns(2)
        kwac = ca.number_input("kWac", 0, 99999, 100, key=f"{key_prefix}_bkwa")
        kwh = cb.number_input("kWh", 0, 99999, 200, key=f"{key_prefix}_bkwh")
        gens.append({"type": "BESS", "kWac": kwac, "cap_val": kwh})
        
    return gens, has_emsb

def get_mv_gen_inputs(key_prefix):
    st.markdown("**MV Generation Source**")
    gen_type = st.radio("Source Type", ["Solar PV", "BESS"], horizontal=True, key=f"{key_prefix}_type")
    
    gens = []
    c1, c2 = st.columns(2)
    
    if gen_type == "Solar PV":
        kwac = c1.number_input("kWac", 0, 99999, 1000, key=f"{key_prefix}_mv_skwa")
        kwp = c2.number_input("kWp", 0, 99999, 1200, key=f"{key_prefix}_mv_skwp")
        gens.append({"type": "Solar", "kWac": kwac, "cap_val": kwp})
    else:
        kwac = c1.number_input("kWac", 0, 99999, 1000, key=f"{key_prefix}_mv_bkwa")
        kwh = c2.number_input("kWh", 0, 99999, 2000, key=f"{key_prefix}_mv_bkwh")
        gens.append({"type": "BESS", "kWac": kwac, "cap_val": kwh})
        
    return gens

def get_feeder_width_config(is_pptx=True):
    # Returns raw float values (inches)
    return {
        "item_w": 3.5,  
        "min_w": 5.0,   
        "gap": 1.0,     
        "sub_gap": 1.0  
    }

def calculate_single_feeder_width(config, dims):
    c_type = config.get("type", "Standard")
    
    if c_type == "Sub-Board":
        sub_feeders = config.get("sub_feeders", {})
        n_subs = len(sub_feeders)
        if n_subs == 0: 
            return dims["min_w"], []
            
        sub_widths = []
        for j in range(n_subs):
            s_conf = sub_feeders.get(j, {})
            # Determine width based on type of sub-feeder
            sf_type = s_conf.get("type", "Standard")
            
            if sf_type == "MV Gen":
                calc_w = dims["item_w"]
            elif sf_type == "Extension":
                calc_w = dims["item_w"]
            else:
                # Standard means LV items
                s_gens = s_conf.get("gens", [])
                s_emsb = s_conf.get("has_emsb", False)
                item_count = len(s_gens) + (1 if s_emsb else 0)
                calc_w = max(dims["item_w"] * 1.5, item_count * dims["item_w"])
            
            sub_widths.append(calc_w)
            
        total_sub_width = sum(sub_widths) + (len(sub_widths) - 1) * dims["sub_gap"]
        final_total_w = max(dims["min_w"], total_sub_width)
        return final_total_w, sub_widths
        
    else:
        gens = config.get("gens", [])
        has_emsb = config.get("emsb", {}).get("has", False)
        item_count = len(gens) + (1 if has_emsb else 0)
        
        if item_count <= 1:
            return dims["min_w"], []
        else:
            return max(dims["min_w"], item_count * dims["item_w"]), []

def calculate_section_layout(section_feeders, swg_configs, start_x, is_pptx=False):
    dims = get_feeder_width_config(is_pptx)
    current_x = start_x
    feeder_centers = []
    feeder_widths = []
    sub_widths_map = {}
    
    if not section_feeders: return 0, [], [], {}

    for i in section_feeders:
        config = swg_configs.get(i, {})
        w, sub_ws = calculate_single_feeder_width(config, dims)
        
        center = current_x + (w / 2)
        feeder_centers.append(center)
        feeder_widths.append(w)
        sub_widths_map[i] = sub_ws
        
        current_x += w + dims["gap"]
        
    total_width = current_x - start_x - dims["gap"] 
    return total_width, feeder_centers, feeder_widths, sub_widths_map

# ============================================================
# 2. PPTX DRAWING HELPERS
# ============================================================

def S(val, scale):
    """Scales a value and returns PPTX Inches object"""
    return Inches(val * scale)

def SPt(val, scale):
    """Scales font size, clamping to minimum readable size"""
    new_size = val * scale
    return Pt(max(6, new_size))

def add_line(slide, x1, y1, x2, y2, width_pt=3, color=RGBColor(0, 112, 192)):
    conn = slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT, int(x1), int(y1), int(x2), int(y2))
    conn.line.width = Pt(width_pt)
    conn.line.color.rgb = color
    return conn

def add_busbar(slide, left, top, width):
    if width <= 0: return
    bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, int(left), int(top), int(width), Inches(0.2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(0, 112, 192)
    bar.line.fill.background()

def add_breaker_x(slide, cx, cy, scale, size_base=0.25, color=RGBColor(0, 112, 192)):
    half = S(size_base, scale)
    add_line(slide, int(cx - half), int(cy - half), int(cx + half), int(cy + half), 3.0, color)
    add_line(slide, int(cx - half), int(cy + half), int(cx + half), int(cy - half), 3.0, color)

def pptx_add_transformer(slide, cx_int, center_y, ratio_txt, tx_id, scale):
    r = S(0.35, scale); d = r * 2
    top_y = center_y - r; bot_y = center_y + r
    
    s1 = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, int(cx_int - r), int(top_y - r), int(d), int(d))
    s1.fill.solid(); s1.fill.fore_color.rgb = RGBColor(255, 255, 255) 
    s1.line.width = Pt(2.0); s1.line.color.rgb = RGBColor(0,0,0)
    
    s2 = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, int(cx_int - r), int(bot_y - r), int(d), int(d))
    s2.fill.solid(); s2.fill.fore_color.rgb = RGBColor(255, 255, 255) 
    s2.line.width = Pt(2.0); s2.line.color.rgb = RGBColor(0,0,0)
    
    add_line(slide, cx_int, int(center_y - S(0.9, scale)), cx_int, int(top_y - r + S(0.05, scale)), 3, RGBColor(0,0,0))
    add_line(slide, cx_int, int(bot_y + r - S(0.05, scale)), cx_int, int(center_y + S(0.9, scale)), 3, RGBColor(0,0,0))

    tb = slide.shapes.add_textbox(int(cx_int + S(0.4, scale)), int(center_y - S(0.8, scale)), S(4.0, scale), S(1.5, scale))
    p = tb.text_frame.paragraphs[0]; p.text = f"{tx_id}\n{ratio_txt}"; 
    p.font.bold = True; p.font.size = Pt(20) # FIXED 20pt

def pptx_add_inverter_branch(slide, cx, start_y, gens, scale):
    if not gens: return
    box_top = start_y + S(1.5, scale)
    add_line(slide, cx, start_y, cx, int(box_top), 2, RGBColor(0, 176, 80))
    gen = gens[0]
    w = S(2.2, scale); h = S(1.5, scale); left = int(cx - w/2)
    
    s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, int(box_top), int(w), int(h))
    s.fill.background(); s.line.color.rgb = RGBColor(0, 176, 80); s.line.width = Pt(2.0)
    add_line(slide, left, int(box_top), int(left+w), int(box_top+h), 2, RGBColor(0, 176, 80))
    add_line(slide, left, int(box_top+h), int(left+w), int(box_top), 2, RGBColor(0, 176, 80))
    
    title = "BESS" if gen['type'] == "BESS" else "SOLAR PV"
    cap_unit = "kWh" if gen['type'] == "BESS" else "kWp"
    lines = [title, f"{gen['kWac']} kWac", f"{gen['cap_val']} {cap_unit}"]
    
    tb = slide.shapes.add_textbox(int(cx - S(3, scale)), int(box_top + h + S(0.1, scale)), S(6, scale), S(2.5, scale))
    tf = tb.text_frame
    for l in lines:
        p = tf.add_paragraph(); p.text = l 
        p.font.size = Pt(20); # FIXED 20pt
        p.font.color.rgb = RGBColor(0, 176, 80); 
        p.alignment = PP_ALIGN.CENTER; p.font.bold = True

def pptx_add_lv_system(slide, cx, start_y, gens, has_emsb, emsb_name, scale):
    if not gens and not has_emsb: return
    items = []
    for g in gens: items.append(('GEN', g))
    if has_emsb: items.append(('EMSB', emsb_name))
    num_items = len(items)
    if num_items == 0: return
    
    spacing = S(4.0, scale)
    total_width = (num_items - 1) * spacing
    start_x_offset = cx - int(total_width / 2)
    
    for idx, (itype, data) in enumerate(items):
        px = start_x_offset + int(idx * spacing)
        box_top = start_y + S(1.5, scale)
        
        add_line(slide, px, int(start_y), px, int(box_top + S(0.05, scale)), 2, RGBColor(0, 176, 80))
        
        if itype == 'GEN':
            w = S(2.2, scale); h = S(1.5, scale); left = int(px - w/2)
            s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, int(box_top), int(w), int(h))
            s.fill.background(); s.line.color.rgb = RGBColor(0, 176, 80); s.line.width = Pt(2.0)
            add_line(slide, left, int(box_top + h), int(left + w), int(box_top), 2, RGBColor(0, 176, 80))
            
            title = "BESS" if data['type'] == "BESS" else "SOLAR PV"
            lines = [title, f"{data['kWac']} kWac", f"{data['cap_val']} {'kWh' if data['type']=='BESS' else 'kWp'}"]
            
            # Increased width and size for better visibility
            tb = slide.shapes.add_textbox(int(px - S(2.5, scale)), int(box_top + h + S(0.1, scale)), S(5.0, scale), S(2.0, scale))
            tf = tb.text_frame
            for l in lines:
                p = tf.add_paragraph(); p.text = l; 
                p.font.size = Pt(20); # FIXED 20pt
                p.font.color.rgb = RGBColor(0, 176, 80); 
                p.alignment = PP_ALIGN.CENTER; p.font.bold = True
                
        elif itype == 'EMSB':
            breaker_y = int(start_y + S(0.8, scale))
            add_breaker_x(slide, px, breaker_y, scale, 0.20)
            w = S(1.6, scale); h = S(0.8, scale); left = int(px - w/2)
            s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, int(box_top), int(w), int(h))
            s.fill.solid(); s.fill.fore_color.rgb = RGBColor(0, 112, 192); s.line.fill.background()
            
            tb = slide.shapes.add_textbox(int(px - S(1.5, scale)), int(box_top + h), S(3, scale), S(0.8, scale))
            tb.text_frame.text = data
            tb.text_frame.paragraphs[0].font.size = Pt(20); # FIXED 20pt
            tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER 

def add_continuation_arrow(slide, x, y, direction, label, scale):
    w_arrow = S(0.5, scale)
    h_arrow = S(0.2, scale)
    y_pos = int(y - S(0.1, scale))
    
    if direction == "next":
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.CHEVRON, int(x), y_pos, w_arrow, h_arrow)
        text_x = x - S(1.5, scale)
        align = PP_ALIGN.RIGHT
    else:
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.CHEVRON, int(x - w_arrow), y_pos, w_arrow, h_arrow)
        shape.rotation = 180
        text_x = x + w_arrow
        align = PP_ALIGN.LEFT

    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(255, 0, 0); shape.line.fill.background()

    if label:
        # Move "To Sheet 2" text ABOVE the line by using y - S(1.0) instead of y - S(0.4)
        tb = slide.shapes.add_textbox(int(text_x), int(y - S(1.0, scale)), S(2.5, scale), S(0.8, scale))
        p = tb.text_frame.paragraphs[0]; p.text = label
        p.font.size = Pt(20); # FIXED 20pt
        p.font.bold = True; p.font.color.rgb = RGBColor(255, 0, 0); p.alignment = align

# ============================================================
# 3. MATPLOTLIB PREVIEW HELPERS
# ============================================================

def draw_tx_mpl(ax, x, y, label, ratio):
    r = 0.3
    c1 = plt.Circle((x, y + 0.25), r, fill=True, fc='white', ec='black', lw=2, zorder=10)
    c2 = plt.Circle((x, y - 0.25), r, fill=True, fc='white', ec='black', lw=2, zorder=10)
    ax.add_patch(c1); ax.add_patch(c2)
    ax.text(x + 0.35, y, f"{label}\n{ratio}", va="center", fontsize=8, fontweight='bold')

def draw_lv_system_mpl(ax, cx, start_y, gens, has_emsb, emsb_name):
    if not gens and not has_emsb: return
    items = []
    for g in gens: items.append(('GEN', g))
    if has_emsb: items.append(('EMSB', emsb_name))
    num = len(items)
    
    spacing = 3.5 
    total_w = (num - 1) * spacing
    start_x = cx - total_w / 2
    
    for idx, (itype, data) in enumerate(items):
        px = start_x + idx * spacing
        box_top = start_y - 1.5
        
        ax.plot([px, px], [start_y, box_top], color="tab:green", lw=3)
        
        if itype == 'GEN':
            rect = patches.Rectangle((px - 0.75, box_top - 1.0), 1.5, 1.0, fill=False, edgecolor="tab:green", lw=2)
            ax.add_patch(rect)
            ax.plot([px - 0.75, px + 0.75], [box_top - 1.0, box_top], color="tab:green", lw=1.5)
            
            title = "BESS" if data['type'] == "BESS" else "SOLAR PV"
            unit = "kWh" if data['type'] == "BESS" else "kWp"
            lines = [title, f"{data['kWac']} kWac", f"{data['cap_val']} {unit}"]
            
            for i, l in enumerate(lines):
                ax.text(px, box_top - 1.2 - (i*0.3), l, color="tab:green", ha="center", fontsize=8, fontweight='bold')
                
        elif itype == 'EMSB':
            by = start_y - 0.8
            ax.plot([px-0.15, px+0.15], [by-0.15, by+0.15], color="tab:blue", lw=2)
            ax.plot([px-0.15, px+0.15], [by+0.15, by-0.15], color="tab:blue", lw=2)
            
            rect = patches.Rectangle((px - 0.6, box_top - 0.6), 1.2, 0.6, fill=True, facecolor="tab:blue", edgecolor="tab:blue")
            ax.add_patch(rect)
            ax.text(px, box_top - 0.9, data, ha="center", va="top", fontsize=8)

def draw_section_feeders_mpl(ax, feeder_indices, x_centers, sub_widths_map, Y_MAIN_BUS, Y_FDR_BRK, swg_configs, swg_names, voltage):
    lv_bus_y = {}; lv_bus_x = {}; lv_bus_edges = {}; sub_feeder_bus_edges = {}
    sub_feeder_lv_coords = {}
    dims = get_feeder_width_config(is_pptx=False)

    for i, cx in zip(feeder_indices, x_centers):
        config = swg_configs.get(i, {})
        ctype = config.get("type", "Standard")
        color = "tab:green" if ctype == "MV Gen" else "tab:blue"
        
        ax.plot([cx, cx], [Y_MAIN_BUS, Y_FDR_BRK], color=color, lw=3)
        ax.plot([cx-0.2, cx+0.2], [Y_FDR_BRK-0.2, Y_FDR_BRK+0.2], color=color, lw=3)
        ax.plot([cx-0.2, cx+0.2], [Y_FDR_BRK+0.2, Y_FDR_BRK-0.2], color=color, lw=3)
        
        if voltage != "400V":
            ax.text(cx, Y_FDR_BRK+0.5, swg_names[i], ha="center", fontsize=10, fontweight="bold")
            
        current_y = Y_FDR_BRK - 0.2
        
        if ctype == "MV Gen":
            gens = config.get("gens", [])
            draw_lv_system_mpl(ax, cx, current_y, gens, False, "")
            
        elif ctype == "Sub-Board":
            sub_voltage = config.get('sub_voltage')
            is_extension = (voltage == sub_voltage)
            
            if not is_extension:
                y_bus_connection = current_y - 1.5
                # Draw Transformer
                ax.plot([cx, cx], [current_y, y_bus_connection + 0.6], color="tab:blue", lw=3)
                draw_tx_mpl(ax, cx, y_bus_connection, f"TX-{i+1}", f"{voltage}/{sub_voltage}")
                
                # DRAW THE MISSING BREAKER FOR SUB-BOARD INCOMER
                y_mv_breaker_main = y_bus_connection - 0.7 
                ax.plot([cx, cx], [y_bus_connection - 0.15, y_mv_breaker_main + 0.2], color="tab:blue", lw=3)
                ax.plot([cx-0.15, cx+0.15], [y_mv_breaker_main-0.15, y_mv_breaker_main+0.15], color="tab:blue", lw=3) 
                ax.plot([cx-0.15, cx+0.15], [y_mv_breaker_main+0.15, y_mv_breaker_main-0.15], color="tab:blue", lw=3) 
                
                y_sub_bus = y_mv_breaker_main - 0.6
                ax.plot([cx, cx], [y_mv_breaker_main - 0.15, y_sub_bus], color="tab:blue", lw=3)
            else:
                # Extension: Direct Line but with a Breaker (Incomer to extension)
                y_ext_brk = current_y - 1.5 
                ax.plot([cx, cx], [current_y, y_ext_brk + 0.2], color="tab:blue", lw=3)
                ax.plot([cx-0.2, cx+0.2], [y_ext_brk-0.2, y_ext_brk+0.2], color="tab:blue", lw=3)
                ax.plot([cx-0.2, cx+0.2], [y_ext_brk+0.2, y_ext_brk-0.2], color="tab:blue", lw=3)
                y_sub_bus = y_ext_brk - 0.8
                ax.plot([cx, cx], [y_ext_brk - 0.2, y_sub_bus], color="tab:blue", lw=3)

            # Sub Bus Bar
            sub_feeders = config.get("sub_feeders", {})
            n_subs = len(sub_feeders)
            widths_list = sub_widths_map.get(i, [])
            
            if n_subs > 0:
                total_sb_width = sum(widths_list) + (len(widths_list)-1)*dims["sub_gap"]
                start_sub_x = cx - (total_sb_width / 2)
                
                ax.hlines(y_sub_bus, start_sub_x, start_sub_x + total_sb_width, lw=4, color="tab:blue")
                ax.text(start_sub_x + total_sb_width + 0.2, y_sub_bus, sub_voltage, color="tab:blue", fontsize=8, fontweight='bold', va='center', ha='left')
                
                current_sb_x = start_sub_x
                for j in range(n_subs):
                    sub_w = widths_list[j]
                    sub_x = current_sb_x + (sub_w / 2)
                    
                    s_conf = sub_feeders.get(j, {})
                    sf_type = s_conf.get("type", "Standard")
                    s_gens = s_conf.get("gens", [])
                    has_emsb = s_conf.get("has_emsb", False)
                    
                    y_sub_mv_brk = y_sub_bus - 0.8
                    ax.plot([sub_x, sub_x], [y_sub_bus, y_sub_mv_brk+0.15], color="tab:blue", lw=2)
                    ax.plot([sub_x-0.1, sub_x+0.1], [y_sub_mv_brk-0.1, y_sub_mv_brk+0.1], color="tab:blue", lw=2)
                    ax.plot([sub_x-0.1, sub_x+0.1], [y_sub_mv_brk+0.1, y_sub_mv_brk-0.1], color="tab:blue", lw=2)
                    
                    if sf_type == "MV Gen":
                        # Direct to Gen (No TX)
                        draw_lv_system_mpl(ax, sub_x, y_sub_mv_brk - 0.2, s_gens, False, "")
                    elif sf_type == "Extension":
                        # Extension from Sub-Board (Same Voltage)
                        # Just a line going down and a label
                        ax.plot([sub_x, sub_x], [y_sub_mv_brk-0.15, y_sub_mv_brk-1.5], color="tab:blue", lw=2)
                        ax.text(sub_x, y_sub_mv_brk - 1.7, f"{sub_voltage} OUT", ha="center", fontsize=7, fontweight="bold")
                    else:
                        # Standard (Implies Step Down TX to 400V)
                        y_tx_sub = y_sub_mv_brk - 1.2
                        ax.plot([sub_x, sub_x], [y_sub_mv_brk-0.15, y_tx_sub+0.3], color="tab:blue", lw=2)
                        draw_tx_mpl(ax, sub_x, y_tx_sub, f"TX-SF{j+1}", f"{sub_voltage}/0.4")
                        
                        y_sub_breaker = y_tx_sub - 1.5
                        ax.plot([sub_x, sub_x], [y_tx_sub - 0.6, y_sub_breaker], color="tab:blue", lw=2)
                        
                        bx, by = sub_x, y_sub_breaker
                        ax.plot([bx-0.1, bx+0.1], [by-0.1, by+0.1], color="tab:blue", lw=2)
                        ax.plot([bx-0.1, bx+0.1], [by+0.1, by-0.1], color="tab:blue", lw=2)
                        
                        y_lv_out = by - 0.2
                        sub_feeder_lv_coords[(i, j)] = (sub_x, y_lv_out)
                        
                        bus_width_viz = max(dims["item_w"], sub_w - 0.5)
                        bus_left = sub_x - bus_width_viz/2
                        bus_right = sub_x + bus_width_viz/2
                        ax.hlines(y_lv_out, bus_left, bus_right, lw=4, color="tab:blue")
                        sub_feeder_bus_edges[(i, j)] = (bus_left, bus_right)
                        draw_lv_system_mpl(ax, sub_x, y_lv_out, s_gens, has_emsb, "EMSB")
                    
                    if sf_type != "Extension":
                        ax.text(sub_x, y_sub_mv_brk - 0.4 if sf_type=="MV Gen" else y_sub_breaker - 0.4, s_conf.get("name", ""), ha="center", fontsize=7)
                    else:
                        ax.text(sub_x, y_sub_mv_brk - 0.4, s_conf.get("name", ""), ha="center", fontsize=7)
                    
                    current_sb_x += sub_w + dims["sub_gap"]

        else: # Standard
            chain = get_tx_chain(voltage, config.get("tx_scheme", ""))
            if not chain and voltage == "400V":
                y_bus = 2.0; ax.plot([cx, cx], [current_y, y_bus], color="tab:blue", lw=3)
                lv_bus_y[i] = y_bus; lv_bus_x[i] = cx
            else:
                for step in chain:
                    y_tx = current_y - 1.5
                    draw_tx_mpl(ax, cx, y_tx, f"TX-{i+1}", step["ratio"])
                    ax.plot([cx, cx], [current_y, y_tx+0.6], color="tab:blue", lw=3, zorder=1)
                    current_y = y_tx - 0.6
                y_bus = current_y - 1.0
                ax.plot([cx, cx], [current_y, y_bus], color="tab:blue", lw=3)
                lv_bus_y[i] = y_bus; lv_bus_x[i] = cx
            
            gens = config.get("gens", [])
            has_emsb = config.get("emsb", {}).get("has")
            
            item_count = len(gens) + (1 if has_emsb else 0)
            bw = max(dims["min_w"], item_count * dims["item_w"]) 
            
            left_edge = cx - bw/2; right_edge = cx + bw/2
            ax.hlines(lv_bus_y[i], left_edge, right_edge, lw=5, color="tab:blue")
            lv_bus_edges[i] = (left_edge, right_edge)
            ax.text(cx, lv_bus_y[i]-0.3, config.get("msb_name", ""), ha="center", va="top", fontweight="bold")
            draw_lv_system_mpl(ax, cx, lv_bus_y[i], gens, has_emsb, config["emsb"]["name"])

    return lv_bus_x, lv_bus_y, lv_bus_edges, sub_feeder_lv_coords, sub_feeder_bus_edges

def draw_preview_mpl(voltage, num_in, num_swg, section_distribution, inc_bc_status, 
                     msb_bc_status, lv_couplers, lv_bc_status, swg_names, swg_configs):
    
    start_margin = 2.0
    total_w = 0
    temp_x = start_margin
    for count in section_distribution:
        if count == 0: continue 
        w, _, _, _ = calculate_section_layout(range(count), swg_configs, temp_x, is_pptx=False)
        total_w += w + 3.0
        temp_x += w + 3.0
        
    fig_w = max(12.0, total_w + 15.0) 
    fig, ax = plt.subplots(figsize=(fig_w, 16))
    
    Y_INC_TOP = 13; Y_INC_BRK = 12; Y_MAIN_BUS = 11; Y_FDR_BRK = 10
    
    current_feeder_idx = 0
    current_x_start = start_margin
    
    global_lv_x = {}; global_lv_y = {}; global_lv_edges = {}; global_sub_coords = {}
    global_sub_bus_edges = {}

    bus_endpoints = []
    
    for s_idx, count in enumerate(section_distribution):
        if count == 0: 
            bus_endpoints.append((None, None)); continue

        indices = list(range(current_feeder_idx, current_feeder_idx + count))
        width, centers, _, sub_w_map = calculate_section_layout(indices, swg_configs, current_x_start, is_pptx=False)
        
        bus_min = min(centers) - 3.0; bus_max = max(centers) + 3.0
        ax.hlines(Y_MAIN_BUS, bus_min, bus_max, lw=6, color="tab:blue")
        bus_endpoints.append((bus_min, bus_max))
        
        inc_x = (bus_min + bus_max) / 2 if num_in == 1 else (bus_max - 1.5 if s_idx == 0 else bus_min + 1.5)
        ax.plot([inc_x, inc_x], [Y_INC_TOP, Y_MAIN_BUS], color="tab:blue", lw=3)
        ax.plot([inc_x-0.2, inc_x+0.2], [Y_INC_BRK-0.2, Y_INC_BRK+0.2], color="tab:blue", lw=3)
        ax.plot([inc_x-0.2, inc_x+0.2], [Y_INC_BRK+0.2, Y_INC_BRK-0.2], color="tab:blue", lw=3)
        ax.text(inc_x, Y_INC_TOP + 0.2, f"INCOMING {s_idx+1}\n({voltage})", ha="center", fontweight="bold")
        
        if s_idx == len(section_distribution) - 1:
             ax.text(bus_max + 0.5, Y_MAIN_BUS, voltage, ha="left", va="center", fontweight="bold", fontsize=12, color="tab:blue")
        
        lv_x, lv_y, lv_edges, sub_c, sub_edges = draw_section_feeders_mpl(ax, indices, centers, sub_w_map, Y_MAIN_BUS, Y_FDR_BRK, swg_configs, swg_names, voltage)
        global_lv_x.update(lv_x); global_lv_y.update(lv_y); global_lv_edges.update(lv_edges)
        global_sub_coords.update(sub_c); global_sub_bus_edges.update(sub_edges)
        
        current_feeder_idx += count
        current_x_start += width + 3.0
        
    for s_idx in range(len(section_distribution) - 1):
        left_sect = bus_endpoints[s_idx]
        right_sect = bus_endpoints[s_idx+1]
        if left_sect[1] is not None and right_sect[0] is not None:
            ax.plot([left_sect[1], right_sect[0]], [Y_MAIN_BUS, Y_MAIN_BUS], color="tab:red", lw=3)
            mid_c = (left_sect[1] + right_sect[0]) / 2
            ax.plot([mid_c-0.2, mid_c+0.2], [Y_MAIN_BUS-0.2, Y_MAIN_BUS+0.2], color="tab:red", lw=3)
            ax.plot([mid_c-0.2, mid_c+0.2], [Y_MAIN_BUS+0.2, Y_MAIN_BUS-0.2], color="tab:red", lw=3)
            ax.text(mid_c, Y_MAIN_BUS + 0.5, f"BC-{s_idx+1}\n({msb_bc_status.get(s_idx, 'NO')})", color="tab:red", ha="center", fontweight="bold")

    for idx in lv_couplers:
        if idx in global_lv_edges and (idx+1) in global_lv_edges:
            y_c = global_lv_y[idx]
            edge_left = global_lv_edges[idx][1]; edge_right = global_lv_edges[idx+1][0]
            ax.plot([edge_left, edge_right], [y_c, y_c], color="tab:red", lw=3)
            mid = (edge_left + edge_right) / 2
            ax.plot([mid-0.2, mid+0.2], [y_c-0.2, y_c+0.2], color="tab:red", lw=3)
            ax.plot([mid-0.2, mid+0.2], [y_c+0.2, y_c-0.2], color="tab:red", lw=3)
            ax.text(mid, y_c+0.5, f"LV-BC\n({lv_bc_status.get(idx, 'NO')})", color="tab:red", ha="center", fontweight="bold", fontsize=9)

    for f_idx, config in swg_configs.items():
        if config.get("type") == "Sub-Board":
            sf_couplers = config.get("sub_couplers", [])
            for j in sf_couplers:
                if (f_idx, j) in global_sub_bus_edges and (f_idx, j+1) in global_sub_bus_edges:
                    _, right_edge_1 = global_sub_bus_edges[(f_idx, j)]
                    left_edge_2, _ = global_sub_bus_edges[(f_idx, j+1)]
                    _, y1 = global_sub_coords[(f_idx, j)]
                    ax.plot([right_edge_1, left_edge_2], [y1, y1], color="tab:red", lw=3)
                    mid = (right_edge_1 + left_edge_2)/2
                    ax.plot([mid-0.15, mid+0.15], [y1-0.15, y1+0.15], color="tab:red", lw=2)
                    ax.plot([mid-0.15, mid+0.15], [y1+0.15, y1-0.15], color="tab:red", lw=2)
                    ax.text(mid, y1+0.3, "LV-BC", color="tab:red", fontsize=6, ha="center")

    if global_lv_x:
        last_idx = max(global_lv_x.keys())
        last_y = global_lv_y[last_idx]; right_most_edge = global_lv_edges[last_idx][1]
        conf = swg_configs.get(last_idx, {})
        lv_label = "400V"
        if conf.get("type") == "Sub-Board": lv_label = conf.get("sub_voltage", "400V")
        
        # Don't label extension lines as 400V if they are HV
        if conf.get("type") == "Sub-Board" and conf.get("sub_voltage") == voltage:
             pass 
        else:
             ax.text(right_most_edge + 0.5, last_y, lv_label, ha="left", va="center", fontweight="bold", fontsize=10, color="tab:blue")

    ax.axis('off'); ax.set_ylim(-8, 14); ax.set_xlim(0, fig_w)
    return fig

# ============================================================
# 4. UPDATED GENERATION LOGIC WITH SCALING
# ============================================================

def draw_feeder_group_on_slide(slide, voltage, feeders_list, swg_configs, swg_names, 
                               start_x, dims, incomer_data, 
                               draw_bc_start, draw_bc_end, bc_label, 
                               scale_factor):
    """
    Draws a group of feeders. All dimensions are raw inches * scale_factor.
    """
    
    # Y Coordinates (Scaled)
    Y_MAIN_BUS = S(6.0, scale_factor)
    Y_INC_TOP = S(1.0, scale_factor)
    Y_INC_BRK = S(4.0, scale_factor)
    Y_FDR_BRK = S(7.5, scale_factor)
    GAP = S(dims["gap"], scale_factor)
    
    current_x = start_x
    
    # Pre-calc widths for precise placement
    feeder_widths = []
    total_group_width = 0
    for idx in feeders_list:
        conf = swg_configs.get(idx, {})
        w_raw, sub_ws_raw = calculate_single_feeder_width(conf, dims)
        w_scaled = S(w_raw, scale_factor)
        # Scale sub-widths
        sub_ws_scaled = [S(sw, scale_factor) for sw in sub_ws_raw]
        
        feeder_widths.append((w_scaled, sub_ws_scaled))
        total_group_width += w_scaled + GAP
        
    # --- DRAW BUSBAR ---
    bus_left = start_x
    bus_right = start_x + total_group_width
    
    if draw_bc_start:
        add_continuation_arrow(slide, bus_left, Y_MAIN_BUS, "prev", "From Sheet 1", scale_factor)
        bus_left += S(0.8, scale_factor)
        add_line(slide, bus_left, Y_MAIN_BUS, bus_left + S(0.8, scale_factor), Y_MAIN_BUS, 3, RGBColor(255,0,0))
        bus_left += S(0.8, scale_factor)
        
    actual_bus_end = bus_left + total_group_width - GAP
    
    if draw_bc_end:
        add_busbar(slide, bus_left, Y_MAIN_BUS, actual_bus_end - bus_left)
        c_start = actual_bus_end
        
        # ----------------------------------------------------
        # UPDATED SHEET 1 -> SHEET 2 TRANSITION LOGIC
        # ----------------------------------------------------
        
        # 1. Extend red busbar significantly to make room
        ext_len = S(2.5, scale_factor)
        c_end = c_start + ext_len
        add_line(slide, c_start, Y_MAIN_BUS, c_end, Y_MAIN_BUS, 3, RGBColor(255,0,0))
        
        # 2. Draw Breaker X centered in extension
        mid_bc = int((c_start + c_end) / 2)
        add_breaker_x(slide, mid_bc, Y_MAIN_BUS, scale_factor, 0.25, RGBColor(255,0,0))
        
        # 3. Text Label: Centered and Moved Higher (Y - 1.5) to align well
        tb = slide.shapes.add_textbox(mid_bc - S(1.5, scale_factor), Y_MAIN_BUS - S(1.5, scale_factor), S(3.0, scale_factor), S(1.2, scale_factor))
        tb.text_frame.text = bc_label
        
        # Apply style to all paragraphs (fixes alignment for both lines)
        for p in tb.text_frame.paragraphs:
            p.font.color.rgb = RGBColor(255,0,0)
            p.font.size = Pt(20) # FIXED 20pt
            p.alignment = PP_ALIGN.CENTER
        
        # 4. Continuation Arrow at far right
        add_continuation_arrow(slide, c_end, Y_MAIN_BUS, "next", "To Sheet 2", scale_factor)
    else:
        add_busbar(slide, bus_left, Y_MAIN_BUS, actual_bus_end - bus_left)

    # --- DRAW FEEDERS ---
    cursor_x = bus_left 
    
    if incomer_data:
        # Improved Incomer Alignment Logic
        inc_x = bus_left + (actual_bus_end - bus_left)/2
        add_line(slide, inc_x, Y_INC_TOP, inc_x, Y_MAIN_BUS)
        add_breaker_x(slide, inc_x, Y_INC_BRK, scale_factor, 0.3)
        
        # Wider textbox, perfectly centered
        tb_w = S(6.0, scale_factor) 
        tb_x = int(inc_x - tb_w / 2)
        tb_y = Y_INC_TOP - S(1.0, scale_factor)
        
        tb = slide.shapes.add_textbox(tb_x, tb_y, tb_w, S(1.5, scale_factor))
        tb.text_frame.text = incomer_data["label"]
        for p in tb.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(20) # FIXED 20pt

    lv_coords_local = {}

    for i, idx in enumerate(feeders_list):
        config = swg_configs.get(idx, {})
        w_feeder_scaled, sub_ws_scaled = feeder_widths[i]
        
        cx = int(cursor_x + w_feeder_scaled/2)
        
        ctype = config.get("type", "Standard")
        col = RGBColor(0,176,80) if ctype == "MV Gen" else RGBColor(0,112,192)

        add_line(slide, cx, Y_MAIN_BUS, cx, Y_FDR_BRK + S(0.1, scale_factor), 3, col)
        add_breaker_x(slide, cx, Y_FDR_BRK, scale_factor, 0.25, col)
        
        if voltage != "400V":
            tb = slide.shapes.add_textbox(cx - S(2, scale_factor), Y_FDR_BRK - S(0.8, scale_factor), S(4, scale_factor), S(0.8, scale_factor))
            tb.text_frame.text = swg_names[idx]
            p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; p.font.bold = True; p.font.size = Pt(16) # Only Feeder Name is 16pt (as standard)

        cur_y = Y_FDR_BRK + S(0.1, scale_factor)
        y_fin_this_feeder = 0; lv_edges = (0,0)

        if ctype == "MV Gen":
            gens = config.get("gens", [])
            pptx_add_inverter_branch(slide, cx, cur_y, gens, scale_factor)
            
        elif ctype == "Sub-Board":
            sub_voltage = config.get('sub_voltage')
            is_extension = (voltage == sub_voltage)
            
            y_tx1 = cur_y + S(2.5, scale_factor)
            y_sub_bus = y_tx1 + S(2.7, scale_factor) # Default spacing
            
            if not is_extension:
                # Normal Sub-Board with Transformer
                pptx_add_transformer(slide, cx, int(y_tx1), f"{voltage}/{sub_voltage}", f"TX-{idx+1}", scale_factor)
                add_line(slide, cx, int(cur_y), cx, int(y_tx1 - S(0.9, scale_factor)))
                
                # DRAW THE MISSING BREAKER FOR SUB-BOARD INCOMER (PPTX)
                y_mv_breaker_main = y_tx1 + S(1.5, scale_factor)
                add_line(slide, cx, int(y_tx1 + S(0.9, scale_factor)), cx, int(y_mv_breaker_main - S(0.2, scale_factor)))
                add_breaker_x(slide, cx, int(y_mv_breaker_main), scale_factor, 0.25)
                
                y_sub_bus = y_mv_breaker_main + S(1.2, scale_factor)
                add_line(slide, cx, int(y_mv_breaker_main + S(0.2, scale_factor)), cx, int(y_sub_bus + S(0.05, scale_factor)))
            else:
                # Extension: Direct Line with Breaker (Incomer)
                y_ext_breaker = y_tx1
                
                # Line down to breaker
                add_line(slide, cx, int(cur_y), cx, int(y_ext_breaker - S(0.2, scale_factor)))
                
                # The Breaker
                add_breaker_x(slide, cx, int(y_ext_breaker), scale_factor, 0.25)
                
                # Line down to bus
                y_sub_bus = y_ext_breaker + S(1.2, scale_factor)
                add_line(slide, cx, int(y_ext_breaker + S(0.2, scale_factor)), cx, int(y_sub_bus + S(0.05, scale_factor)))
                
            sub_feeders = config.get("sub_feeders", {})
            n_subs = len(sub_feeders)
            if n_subs > 0:
                total_sb_width = sum(sub_ws_scaled) + (len(sub_ws_scaled)-1)*S(dims["sub_gap"], scale_factor)
                start_sub_x = cx - int(total_sb_width / 2)
                add_busbar(slide, start_sub_x, int(y_sub_bus), total_sb_width)
                
                # CHANGED FONT SIZE TO 20pt
                tb = slide.shapes.add_textbox(start_sub_x + total_sb_width, y_sub_bus - S(0.3, scale_factor), S(1.0, scale_factor), S(0.5, scale_factor))
                tb.text_frame.text = sub_voltage; tb.text_frame.paragraphs[0].font.size = Pt(20)

                curr_sb_x = start_sub_x
                sub_bus_edges_local = {}; sub_y_local = {}

                for j in range(n_subs):
                    sw = sub_ws_scaled[j]; sx = curr_sb_x + int(sw/2)
                    s_conf = sub_feeders.get(j, {})
                    sf_type = s_conf.get("type", "Standard")
                    
                    y_mv_brk_sub = y_sub_bus + S(1.2, scale_factor)
                    add_line(slide, sx, int(y_sub_bus), sx, int(y_mv_brk_sub - S(0.2, scale_factor)))
                    add_breaker_x(slide, sx, int(y_mv_brk_sub), scale_factor, 0.2)
                    
                    if sf_type == "MV Gen":
                        # Direct to Gen
                        pptx_add_inverter_branch(slide, sx, int(y_mv_brk_sub + S(0.2, scale_factor)), s_conf.get("gens", []), scale_factor)
                    elif sf_type == "Extension":
                        # Extension from Sub-Board
                        add_line(slide, sx, int(y_mv_brk_sub + S(0.2, scale_factor)), sx, int(y_mv_brk_sub + S(2.5, scale_factor)))
                        
                        tb = slide.shapes.add_textbox(sx - S(1.5, scale_factor), int(y_mv_brk_sub + S(2.7, scale_factor)), S(3.0, scale_factor), S(0.8, scale_factor))
                        tb.text_frame.text = f"{sub_voltage} OUT"
                        tb.text_frame.paragraphs[0].font.size = Pt(14)
                        tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                        tb.text_frame.paragraphs[0].font.bold = True
                    else:
                        # Standard -> Step Down TX
                        y_tx_sub = y_mv_brk_sub + S(2.2, scale_factor)
                        add_line(slide, sx, int(y_mv_brk_sub + S(0.2, scale_factor)), sx, int(y_tx_sub - S(0.9, scale_factor)))
                        pptx_add_transformer(slide, sx, int(y_tx_sub), f"{sub_voltage}/0.4", f"TX-SF{j+1}", scale_factor)
                        
                        y_sub_breaker = y_tx_sub + S(2.0, scale_factor)
                        add_line(slide, sx, int(y_tx_sub + S(0.9, scale_factor)), sx, int(y_sub_breaker))
                        add_breaker_x(slide, sx, int(y_sub_breaker), scale_factor, 0.2)
                        
                        y_lv_out = y_sub_breaker + S(0.2, scale_factor)
                        b_viz = max(S(dims["item_w"], scale_factor), sw - S(0.5, scale_factor))
                        b_start = sx - int(b_viz/2)
                        b_end = b_start + b_viz
                        add_busbar(slide, b_start, int(y_lv_out), b_viz)
                        
                        # Bottom Label for LAST feeder in loop (400V)
                        if j == n_subs - 1 and i == len(feeders_list) - 1:
                            tb = slide.shapes.add_textbox(b_end + S(0.1, scale_factor), y_lv_out - S(0.3, scale_factor), S(1.5, scale_factor), S(0.6, scale_factor))
                            tb.text_frame.text = "400V"
                            p = tb.text_frame.paragraphs[0]
                            p.font.size = Pt(20) # FIXED 20pt
                            p.font.bold = True
                            p.font.color.rgb = RGBColor(0, 112, 192)
                            p.alignment = PP_ALIGN.LEFT
                        
                        sub_bus_edges_local[j] = (b_start, b_start + b_viz)
                        sub_y_local[j] = y_lv_out
                        
                        pptx_add_lv_system(slide, sx, y_lv_out, s_conf.get("gens", []), s_conf.get("has_emsb"), "EMSB", scale_factor)
                    
                    curr_sb_x += sw + S(dims["sub_gap"], scale_factor)
                
                for cp in config.get("sub_couplers", []):
                     if cp in sub_bus_edges_local and (cp+1) in sub_bus_edges_local:
                         e1 = sub_bus_edges_local[cp][1]; e2 = sub_bus_edges_local[cp+1][0]; y_cp = sub_y_local[cp]
                         add_line(slide, e1, y_cp, e2, y_cp, 3, RGBColor(255,0,0))
                         add_breaker_x(slide, int((e1+e2)/2), y_cp, scale_factor, 0.2, RGBColor(255,0,0))

        else: # Standard
            chain = get_tx_chain(voltage, config.get("tx_scheme", ""))
            temp_y = cur_y
            if not chain and voltage == "400V":
                y_fin_this_feeder = S(14.0, scale_factor) 
                add_line(slide, cx, temp_y, cx, y_fin_this_feeder)
            else:
                for step in chain:
                    y_tx = temp_y + S(2.5, scale_factor)
                    pptx_add_transformer(slide, cx, int(y_tx), step["ratio"], f"TX-{idx+1}", scale_factor)
                    add_line(slide, cx, int(temp_y), cx, int(y_tx - S(0.9, scale_factor)))
                    temp_y = y_tx + S(0.9, scale_factor)
                y_fin_this_feeder = temp_y + S(2.0, scale_factor)
                add_line(slide, cx, int(temp_y), cx, int(y_fin_this_feeder + S(0.05, scale_factor)))
            
            gens = config.get("gens", []); has_emsb = config.get("emsb", {}).get("has")
            cnt = len(gens) + (1 if has_emsb else 0)
            bw = max(S(dims["min_w"], scale_factor), cnt * S(dims["item_w"], scale_factor))
            
            left_edge = cx - int(bw/2); right_edge = cx + int(bw/2)
            add_busbar(slide, left_edge, int(y_fin_this_feeder), bw)
            
            # Bottom Label for LAST feeder in loop (400V)
            if i == len(feeders_list) - 1:
                tb = slide.shapes.add_textbox(right_edge + S(0.1, scale_factor), y_fin_this_feeder - S(0.3, scale_factor), S(1.5, scale_factor), S(0.6, scale_factor))
                tb.text_frame.text = "400V"
                p = tb.text_frame.paragraphs[0]
                p.font.size = Pt(20) # FIXED 20pt
                p.font.bold = True
                p.font.color.rgb = RGBColor(0, 112, 192)
                p.alignment = PP_ALIGN.LEFT
            
            lv_edges = (left_edge, right_edge)
            pptx_add_lv_system(slide, cx, int(y_fin_this_feeder), gens, has_emsb, config["emsb"]["name"], scale_factor)

        if ctype == "Standard":
            lv_coords_local[idx] = {"y": y_fin_this_feeder, "left": lv_edges[0], "right": lv_edges[1]}

        cursor_x += w_feeder_scaled + GAP

    drawn_width = cursor_x - start_x
    return lv_coords_local, drawn_width, actual_bus_end


def generate_pptx(voltage, num_in, num_swg, section_distribution, inc_bc_status, 
                  msb_bc_status, lv_couplers, lv_bc_status, swg_names, swg_configs):
    
    prs = Presentation()
    
    # PPTX has a hard limit around 56 inches.
    MAX_PPTX_WIDTH_INCHES = 56.0 
    
    dims = get_feeder_width_config(is_pptx=True)
    GAP_RAW = dims["gap"]
    
    # 1. GROUP FEEDERS
    sections = []
    curr = 0
    for count in section_distribution:
        if count > 0:
            sections.append(list(range(curr, curr + count)))
        curr += count
    
    # 2. CALCULATE RAW WIDTHS (INCHES)
    section_raw_widths = []
    has_sub_board = False
    for sec_indices in sections:
        total_w = 0
        for idx in sec_indices:
            conf = swg_configs.get(idx, {})
            if conf.get("type") == "Sub-Board": has_sub_board = True
            w, _ = calculate_single_feeder_width(conf, dims)
            total_w += w + GAP_RAW
        section_raw_widths.append(total_w)
    
    # 3. DETERMINE SPLIT STRATEGY
    # Add buffer for Bus Couplers
    total_raw_width = sum(section_raw_widths) + (len(sections)-1)*2.0 
    
    # Dynamic Sizing Logic
    # We want the content to fit nicely.
    # Calculate needed width including margins (e.g. 4 inches total)
    needed_width = total_raw_width + 4.0
    
    # Split threshold: if raw width > max width, we split.
    requires_split = needed_width > MAX_PPTX_WIDTH_INCHES

    global_lv_map = {}
    
    # Decide Height based on content
    needed_height = 30.0 if has_sub_board else 20.0
    Y_MAIN_BUS_CONST = 6.0 # Match function default

    # --- SCENARIO A: SINGLE SLIDE ---
    if not requires_split:
        # Scale to maximize usage if needed, or just set slide size
        scale = 1.0
        # If needed width is huge but <56, let's make the slide that big
        # If needed width is small, make it at least 20 inches wide
        final_w_inches = max(20.0, needed_width)
        
        prs.slide_width = int(Inches(final_w_inches))
        prs.slide_height = int(Inches(needed_height))
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        start_margin = Inches(final_w_inches - total_raw_width) / 2
        current_x = start_margin
        
        for s_i, feeders in enumerate(sections):
            inc_label = f"INCOMING {s_i+1}\n({voltage})"
            is_last = (s_i == len(sections) - 1)
            
            coords, w_used, bus_end_x = draw_feeder_group_on_slide(slide, voltage, feeders, swg_configs, swg_names, 
                                                current_x, dims, 
                                                {"label": inc_label}, 
                                                False, False, "", scale)
            
            for idx, data in coords.items(): 
                data['slide'] = slide
                data['scale'] = scale # Store scale for LV coupler font
                global_lv_map[idx] = data
            current_x += w_used
            
            if not is_last:
                # Draw simple coupler
                gap_size = S(1.0, scale)
                # Align coupler red line with main bus y
                y_bus = S(Y_MAIN_BUS_CONST, scale)
                
                # Draw Red Line
                add_line(slide, current_x - S(GAP_RAW, scale), y_bus, current_x + gap_size, y_bus, 3, RGBColor(255,0,0))
                
                mid_x = int(current_x + gap_size/2 - S(GAP_RAW, scale)/2)
                add_breaker_x(slide, mid_x, y_bus, scale, 0.25, RGBColor(255,0,0))
                
                # COUPLER TEXT (Main Bus) - Font Size 20pt
                bc_text = f"BC-{s_i+1}\n({msb_bc_status.get(s_i, 'NO')})"
                tb_bc = slide.shapes.add_textbox(mid_x - S(1.5, scale), y_bus - S(1.2, scale), S(3.0, scale), S(0.8, scale))
                tb_bc.text_frame.text = bc_text
                for p in tb_bc.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
                    p.font.color.rgb = RGBColor(255,0,0)
                    p.font.size = Pt(20) # FIXED 20pt

                current_x += gap_size
            else:
                # This is the "rightest" bar on the single slide. Add Label 20pt.
                tb = slide.shapes.add_textbox(bus_end_x + S(0.1, scale), S(Y_MAIN_BUS_CONST, scale) - S(0.3, scale), S(1.5, scale), S(0.6, scale))
                tb.text_frame.text = voltage
                p = tb.text_frame.paragraphs[0]
                p.font.size = Pt(20) # FIXED 20pt
                p.font.bold = True
                p.font.color.rgb = RGBColor(0, 112, 192)
                p.alignment = PP_ALIGN.LEFT

    # --- SCENARIO B: SPLIT (LHS / RHS) WITH SCALING ---
    else:
        # 1. Calculate Raw Content Widths
        lhs_raw_w = 0
        rhs_raw_w = 0
        lhs_fds = []
        rhs_fds = []
        
        if len(sections) == 1:
            # Single incomer split logic
            all_fds = sections[0]
            mid = len(all_fds)//2
            lhs_fds = all_fds[:mid]
            rhs_fds = all_fds[mid:]
            
            for idx in lhs_fds: lhs_raw_w += calculate_single_feeder_width(swg_configs[idx], dims)[0] + GAP_RAW
            lhs_raw_w += 2.0 # BC buffer
            
            for idx in rhs_fds: rhs_raw_w += calculate_single_feeder_width(swg_configs[idx], dims)[0] + GAP_RAW
            rhs_raw_w += 2.0
            
        else:
            # Multi-incomer split logic
            lhs_fds = sections[0]
            lhs_raw_w = section_raw_widths[0] + 2.0
            
            rhs_indices_groups = sections[1:]
            rhs_raw_w = sum(section_raw_widths[1:]) + (len(rhs_indices_groups)-1)*1.0 + 2.0

        # 2. Determine Optimal Slide Width
        max_content_w = max(lhs_raw_w, rhs_raw_w)
        target_slide_w = max_content_w + 2.0
        
        final_slide_w = min(target_slide_w, MAX_PPTX_WIDTH_INCHES)
        final_slide_w = max(final_slide_w, 20.0) 
        
        prs.slide_width = int(Inches(final_slide_w))
        prs.slide_height = int(Inches(needed_height))
        
        # 3. Calculate Scales (Horizontal AND Vertical check)
        available_w = final_slide_w - 2.0
        
        # Horizontal fit
        scale_lhs_w = min(1.0, available_w / lhs_raw_w)
        scale_rhs_w = min(1.0, available_w / rhs_raw_w)
        
        # Vertical fit assumption
        scale_h_limit = 0.85
        
        scale_lhs = min(scale_lhs_w, scale_h_limit)
        scale_rhs = min(scale_rhs_w, scale_h_limit)
        
        # 4. Draw Slide 1
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        lhs_content_w = S(lhs_raw_w, scale_lhs)
        start_x1 = (Inches(final_slide_w) - lhs_content_w) / 2
        
        if len(sections) == 1:
             c1, _, _ = draw_feeder_group_on_slide(slide1, voltage, lhs_fds, swg_configs, swg_names, 
                                            start_x1, dims, {"label": f"INCOMING 1\n({voltage})"}, 
                                            False, True, "Bus Cont.", scale_lhs)
        else:
             c1, _, _ = draw_feeder_group_on_slide(slide1, voltage, lhs_fds, swg_configs, swg_names, 
                                            start_x1, dims, {"label": f"INCOMING 1\n({voltage})"}, 
                                            False, True, f"BC-1\n({msb_bc_status.get(0, 'NO')})", scale_lhs)
        for idx, data in c1.items(): 
            data['slide'] = slide1
            data['scale'] = scale_lhs
            global_lv_map[idx] = data
        
        # 5. Draw Slide 2
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        rhs_content_w = S(rhs_raw_w, scale_rhs)
        start_x2 = (Inches(final_slide_w) - rhs_content_w) / 2
        
        if len(sections) == 1:
             c2, _, bus_end_x2 = draw_feeder_group_on_slide(slide2, voltage, rhs_fds, swg_configs, swg_names, 
                                            start_x2, dims, {"label": ""}, 
                                            True, False, "", scale_rhs)
             for idx, data in c2.items(): 
                 data['slide'] = slide2
                 data['scale'] = scale_rhs
                 global_lv_map[idx] = data
             
             # Label rightest bar on Slide 2 - Font Size 20pt
             tb = slide2.shapes.add_textbox(bus_end_x2 + S(0.1, scale_rhs), S(Y_MAIN_BUS_CONST, scale_rhs) - S(0.3, scale_rhs), S(1.5, scale_rhs), S(0.6, scale_rhs))
             tb.text_frame.text = voltage
             p = tb.text_frame.paragraphs[0]
             p.font.size = Pt(20) # FIXED 20pt
             p.font.bold = True
             p.font.color.rgb = RGBColor(0, 112, 192)
             p.alignment = PP_ALIGN.LEFT

        else:
             curr_x = start_x2
             rhs_indices_groups = sections[1:]
             for r_i, r_feeders in enumerate(rhs_indices_groups):
                is_first = (r_i == 0)
                # Correct Incoming Label Logic for Slide 2
                real_inc_idx = r_i + 2 # Since incoming 1 is on slide 1
                lbl = f"INCOMING {real_inc_idx}\n({voltage})"
                
                c_out, w_used, bus_end_x = draw_feeder_group_on_slide(slide2, voltage, r_feeders, swg_configs, swg_names, 
                                                    curr_x, dims, {"label": lbl},
                                                    is_first, False, "", scale_rhs)
                for idx, data in c_out.items(): 
                    data['slide'] = slide2
                    data['scale'] = scale_rhs
                    global_lv_map[idx] = data
                
                if r_i < len(rhs_indices_groups) - 1:
                    curr_x += w_used
                    add_line(slide2, curr_x - S(GAP_RAW, scale_rhs), S(6.0, scale_rhs), curr_x + S(1.0, scale_rhs), S(6.0, scale_rhs), 3, RGBColor(255,0,0))
                    add_breaker_x(slide2, curr_x, S(6.0, scale_rhs), scale_rhs, 0.25, RGBColor(255,0,0))
                    
                    # COUPLER TEXT (Internal) - Font Size 20pt
                    bc_text = f"BC-{r_i+2}\n({msb_bc_status.get(r_i+1, 'NO')})"
                    tb_bc = slide2.shapes.add_textbox(curr_x - S(0.5, scale_rhs), S(6.0, scale_rhs) - S(1.2, scale_rhs), S(3.0, scale_rhs), S(0.8, scale_rhs))
                    tb_bc.text_frame.text = bc_text
                    for p in tb_bc.text_frame.paragraphs:
                        p.alignment = PP_ALIGN.CENTER
                        p.font.color.rgb = RGBColor(255,0,0)
                        p.font.size = Pt(20) # FIXED 20pt

                    curr_x += S(1.0, scale_rhs)
                else:
                    # Last group on Slide 2 -> Label rightest bar - Font Size 20pt
                    tb = slide2.shapes.add_textbox(bus_end_x + S(0.1, scale_rhs), S(Y_MAIN_BUS_CONST, scale_rhs) - S(0.3, scale_rhs), S(1.5, scale_rhs), S(0.6, scale_rhs))
                    tb.text_frame.text = voltage
                    p = tb.text_frame.paragraphs[0]
                    p.font.size = Pt(20) # FIXED 20pt
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 112, 192)
                    p.alignment = PP_ALIGN.LEFT

    # 4. LV COUPLERS (Only if same slide)
    for pair_idx in lv_couplers:
        f1 = pair_idx; f2 = pair_idx + 1
        if f1 in global_lv_map and f2 in global_lv_map:
            d1 = global_lv_map[f1]; d2 = global_lv_map[f2]
            if d1['slide'] == d2['slide']:
                sl = d1['slide']
                scale_used = d1.get('scale', 1.0) # Retrieve stored scale
                y_c = int((d1['y'] + d2['y']) / 2)
                x1 = d1['right']; x2 = d2['left']
                
                add_line(sl, x1, y_c, x2, y_c, 3, RGBColor(255,0,0))
                mid = int((x1+x2)/2)
                add_breaker_x(sl, mid, y_c, 1.0, 0.2, RGBColor(255,0,0)) 
                
                # LV Coupler Text - Font Size 20pt
                tb = sl.shapes.add_textbox(mid - Inches(1.5), y_c - Inches(0.8), Inches(3.0), Inches(0.8))
                tb.text_frame.text = f"LV-BC ({lv_bc_status.get(f1, 'NO')})"
                tb.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,0,0)
                tb.text_frame.paragraphs[0].font.size = Pt(20) # FIXED 20pt
                tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()

def main():
    st.set_page_config(layout="wide", page_title="SLD Generator")
    
    # --- AUTHENTICATION LOGIC ---
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.title("🔒 Internal Access Only")
        passcode = st.text_input("Enter Passcode:", type="password")
        if st.button("Login"):
            if passcode == "9999":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect Passcode")
        return # Stop execution here if not authenticated

    # --- MAIN APP LOGIC (Only runs if authenticated) ---
    st.title("⚡ SLD Generator")

    with st.sidebar:
        st.subheader("System Configuration")
        if st.button("Reset All", type="secondary"):
            # Clear all keys except authentication
            for key in list(st.session_state.keys()):
                if key != "authenticated":
                    del st.session_state[key]
            st.rerun()

        voltage = st.selectbox("Voltage", ["400V", "11kV", "33kV", "132kV"], key="sys_v")
        num_in = st.selectbox("Incomers", [1, 2, 3], index=1, key="sys_in")
        n_swg = st.number_input("Total Number of Feeders", 1, 20, 4)
        
        section_distribution = []
        if num_in == 1:
            section_distribution = [n_swg]
        else:
            st.markdown("### Bus Section Configuration")
            remaining = n_swg
            for i in range(num_in - 1):
                val = st.number_input(f"Feeders on Bus Section {i+1}", 0, remaining, max(1, remaining//2), key=f"sec_{i}")
                section_distribution.append(val); remaining -= val
            section_distribution.append(remaining)
            st.info(f"Feeders on Bus Section {num_in}: {remaining}")

        msb_bc_status = {}
        if num_in > 1:
            st.markdown("### Bus Coupler Status")
            for i in range(num_in - 1):
                msb_bc_status[i] = st.selectbox(f"Bus Coupler {i+1}-{i+2}", ["NO", "NC"], key=f"mbc_{i}")

        swg_configs = {}; swg_names = []
        
        st.markdown("### Feeder Details")
        for i in range(n_swg):
            with st.expander(f"Feeder {i+1}", expanded=False):
                name = st.text_input("Name", f"F-{i+1}", key=f"n_{i}")
                swg_names.append(name)
                
                valid_types = ["Standard", "MV Gen"]
                if voltage not in ["400V", "11kV"]: valid_types.append("Sub-Board")
                if voltage != "400V": valid_types.append("Extension")
                
                ctype = "Standard"
                if voltage != "400V": ctype = st.selectbox("Type", valid_types, key=f"t_{i}")
                
                gens = []; has_emsb = False
                
                if ctype == "MV Gen":
                    gens = get_mv_gen_inputs(f"g_{i}")
                elif ctype == "Standard":
                    gens, has_emsb = get_lv_gen_inputs(f"g_{i}", include_emsb=True)
                
                conf = {"type": ctype, "msb_name": name, "gens": gens, "emsb": {"has": has_emsb, "name": "EMSB"}}
                
                if ctype == "Standard":
                    conf["tx_scheme"] = f"{voltage}/0.4 kV"
                    
                elif ctype == "Sub-Board":
                    conf["sub_voltage"] = st.selectbox("Sub Voltage", ["11kV", "6.6kV"], key=f"sv_{i}")
                    n_sub_feeders = st.number_input(f"No. of {conf['sub_voltage']} Feeders", 1, 10, 2, key=f"nsf_{i}")
                    
                    if n_sub_feeders > 1:
                        valid_s_pairs = list(range(n_sub_feeders - 1))
                        s_labels = [f"SF-{p+1} & SF-{p+2}" for p in valid_s_pairs]
                        sel_s_couplers = st.multiselect("Add LV Coupler between:", s_labels, key=f"ssc_{i}")
                        conf["sub_couplers"] = [valid_s_pairs[s_labels.index(l)] for l in sel_s_couplers]
                    
                    conf["sub_feeders"] = {}
                    for j in range(n_sub_feeders):
                        st.caption(f"Sub-Feeder {j+1}")
                        sf_name = st.text_input(f"Name", f"SF-{j+1}", key=f"sfn_{i}_{j}")
                        
                        # Added "Extension" type to Sub-Board sub-feeders
                        sf_type_options = ["Standard", "MV Gen", "Extension"]
                        sf_type = st.selectbox("Sub-Feeder Type", sf_type_options, key=f"sft_{i}_{j}")
                        
                        sf_gens = []
                        sf_emsb = False
                        
                        if sf_type == "Standard":
                            sf_gens, sf_emsb = get_lv_gen_inputs(f"sfg_{i}_{j}", include_emsb=True)
                        elif sf_type == "MV Gen":
                            sf_gens = get_mv_gen_inputs(f"sfg_{i}_{j}")
                        # Extension requires no extra inputs, just Name
                            
                        conf["sub_feeders"][j] = {"type": sf_type, "name": sf_name, "gens": sf_gens, "has_emsb": sf_emsb}
                
                elif ctype == "Extension":
                    # Map Extension to Sub-Board structure but same voltage
                    conf["type"] = "Sub-Board" 
                    conf["sub_voltage"] = voltage
                    
                    st.info(f"Extension at {voltage}")
                    n_sub_feeders = st.number_input(f"No. of Feeders on Extension", 1, 10, 2, key=f"nef_{i}")
                    
                    # Logic to Add Couplers for Extension (similar to Sub-Board)
                    if n_sub_feeders > 1:
                        valid_e_pairs = list(range(n_sub_feeders - 1))
                        e_labels = [f"EF-{p+1} & EF-{p+2}" for p in valid_e_pairs]
                        sel_e_couplers = st.multiselect("Add Bus Coupler between:", e_labels, key=f"ext_bc_{i}")
                        conf["sub_couplers"] = [valid_e_pairs[e_labels.index(l)] for l in sel_e_couplers]
                    
                    conf["sub_feeders"] = {}
                    for j in range(n_sub_feeders):
                        st.caption(f"Extension Feeder {j+1}")
                        sf_name = st.text_input(f"Name", f"EF-{j+1}", key=f"efn_{i}_{j}")
                        
                        sf_type = st.selectbox("Type", ["Standard", "MV Gen"], key=f"eft_{i}_{j}")
                        
                        sf_gens = []
                        sf_emsb = False
                        
                        if sf_type == "Standard":
                             sf_gens, sf_emsb = get_lv_gen_inputs(f"efg_{i}_{j}", include_emsb=True)
                        else:
                             sf_gens = get_mv_gen_inputs(f"efg_{i}_{j}")
                             
                        conf["sub_feeders"][j] = {"type": sf_type, "name": sf_name, "gens": sf_gens, "has_emsb": sf_emsb}

                swg_configs[i] = conf

        lv_couplers = []; lv_bc_status = {}
        with st.expander("LV (0.4kV) Bus Couplers"):
            valid_pairs = []
            for i in range(n_swg-1):
                if swg_configs[i]["type"] == "Standard" and swg_configs[i+1]["type"] == "Standard":
                    valid_pairs.append(i)
            if valid_pairs:
                pair_lbls = [f"#{p+1} & #{p+2}" for p in valid_pairs]
                sel_lv = st.multiselect("Couples", pair_lbls, key="lv_c_sel")
                for s in sel_lv:
                    idx = pair_lbls.index(s); real_idx = valid_pairs[idx]
                    lv_couplers.append(real_idx); lv_bc_status[real_idx] = st.selectbox(f"Status {s}", ["NO", "NC"], key=f"lvbc_{real_idx}")
            else:
                st.write("No adjacent standard feeders available for coupling.")

    st.subheader("Preview")
    fig = draw_preview_mpl(voltage, num_in, n_swg, section_distribution, [], msb_bc_status, 
                           lv_couplers, lv_bc_status, swg_names, swg_configs)
    st.pyplot(fig, use_container_width=True)
    
    pptx_data = generate_pptx(voltage, num_in, n_swg, section_distribution, [], msb_bc_status, 
                              lv_couplers, lv_bc_status, swg_names, swg_configs)
    
    st.download_button("📥 Download PowerPoint", pptx_data, 
                       f"SLD_{voltage}.pptx", 
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                       type="primary", use_container_width=True)

if __name__ == "__main__":
    main()
