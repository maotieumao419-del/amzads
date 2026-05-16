import streamlit as st
import json
import os

# --- Configuration ---
RULES_FILE = r"c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\rules.json"

st.set_page_config(page_title="PPC Rules Manager", layout="wide", page_icon="⚙️")

# Custom CSS for better aesthetics
st.markdown("""
<style>
    .stExpander header p {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1E88E5;
    }
    .stButton>button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem;
    }
    .stButton>button:hover {
        background-color: #45a049;
        border-color: #45a049;
    }
</style>
""", unsafe_allow_html=True)

def load_rules():
    with open(RULES_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_rules(data):
    with open(RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def get_options_for_key(key, current_val):
    options_map = {
        "action": ["Update", "Paused", "Enable"],
        "target_entity": ["Keyword", "Bidding Adjustment", "Campaign", "Ad Group"],
        "placement_type": ["Placement Top", "Placement Product Page", "Placement Rest Of Search"]
    }
    
    # If we have mapping for this key, use it
    if key in options_map:
        opts = options_map[key].copy()
        if current_val and current_val not in opts:
            opts.append(current_val)
        return opts
    
    # If no predefined mapping but it's a text/action key, we return at least the current val
    return [current_val] if current_val else [""]

def render_scalar(widget_key, k, v):
    key_label = k.replace("_", " ").title()
    
    # Booleans
    if isinstance(v, bool):
        return st.checkbox(key_label, value=v, key=widget_key)
        
    # Integers (clicks, impressions, etc.)
    elif isinstance(v, int):
        return st.number_input(key_label, value=v, step=1, key=widget_key)
        
    # Floats (bid_cap, %, etc.)
    elif isinstance(v, float):
        return st.number_input(key_label, value=v, step=0.01, format="%.2f", key=widget_key)
        
    # Strings (text/actions)
    elif isinstance(v, str):
        # Auto-detect if it's a numeric string like "25" for percentages
        if v.isdigit():
            val_int = int(v)
            # Dùng number input nhưng sẽ trả về kiểu string để giữ nguyên cấu trúc
            edited_val = st.number_input(key_label, value=val_int, step=1, key=widget_key)
            return str(edited_val)
            
        # Selectbox for action/entity/placement
        elif k in ["action", "target_entity", "placement_type", "match_type"] or "action" in k:
            options = get_options_for_key(k, v)
            index = options.index(v) if v in options else 0
            return st.selectbox(key_label, options=options, index=index, key=widget_key)
            
        # Text area cho log templates
        elif "template" in k or "log" in k:
            return st.text_area(key_label, value=v, key=widget_key)
            
        # Text input cho các string còn lại (ví dụ rule_name)
        else:
            return st.text_input(key_label, value=v, key=widget_key)
            
    else:
        st.warning(f"Type unsupported: {type(v)} for {k}")
        return v

def render_dict(prefix, d, depth=0):
    edited_d = {}
    
    scalar_keys = [k for k, v in d.items() if not isinstance(v, dict)]
    dict_keys = [k for k, v in d.items() if isinstance(v, dict)]
    
    if scalar_keys:
        cols = st.columns(min(len(scalar_keys), 3) if len(scalar_keys) > 0 else 1)
        for i, k in enumerate(scalar_keys):
            v = d[k]
            col = cols[i % 3]
            with col:
                edited_d[k] = render_scalar(f"{prefix}_{k}", k, v)
                
    if dict_keys:
        for k in dict_keys:
            if depth > 0:
                st.markdown(f"**{k.replace('_', ' ').title()}**")
            else:
                st.markdown(f"##### 🔹 {k.replace('_', ' ').title()}")
                
            with st.container(border=True):
                edited_d[k] = render_dict(f"{prefix}_{k}", d[k], depth + 1)
                
    return edited_d

def validate_rules(rules):
    # Validate no empty strings in string fields
    def is_valid(d):
        for k, v in d.items():
            if isinstance(v, dict):
                if not is_valid(v):
                    return False
            elif isinstance(v, str) and not v.strip():
                return False
        return True
        
    for r in rules:
        if not is_valid(r):
            return False
    return True

def main():
    st.title("🎯 Amazon PPC Automated Trading - Rule Manager")
    st.markdown("Quản lý cấu hình thuật toán tự động từ Nguồn chân lý `rules.json`")
    st.divider()

    if not os.path.exists(RULES_FILE):
        st.error(f"❌ Không tìm thấy file: `{RULES_FILE}`")
        return

    try:
        data = load_rules()
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc JSON: {e}")
        return

    rules = data.get("rules", [])
    if not rules:
        st.warning("⚠️ File cấu hình không có rule nào.")
        return

    edited_rules = []

    # Create the form to batch all changes
    with st.form("save_rules_form"):
        st.markdown("### 📝 Danh sách Rules")
        
        for i, rule in enumerate(rules):
            rule_name = rule.get("rule_name", f"Rule_{i}")
            # Dùng expander để nhóm lại, tránh rối mắt
            with st.expander(f"⚙️ Rule: {rule_name}", expanded=False):
                edited_rule = render_dict(f"rule_{i}", rule)
                edited_rules.append(edited_rule)
                
        st.divider()
        submit_btn = st.form_submit_button("💾 LƯU CẤU HÌNH (SAVE RULES)")
        
        if submit_btn:
            if validate_rules(edited_rules):
                data["rules"] = edited_rules
                try:
                    save_rules(data)
                    st.success("✅ Đã lưu cấu hình `rules.json` thành công! Form đã được validate an toàn.")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ Lỗi ghi file: {e}")
            else:
                st.error("🚨 Dữ liệu không hợp lệ: Các trường văn bản (Text/Action) không được để trống!")

if __name__ == "__main__":
    main()
