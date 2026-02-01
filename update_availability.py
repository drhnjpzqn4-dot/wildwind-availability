#!/usr/bin/env python3
"""
Wildwind Room Availability Updater
Downloads Excel from Dropbox and generates an interactive HTML dashboard.
"""

import pandas as pd
import json
import requests
from datetime import datetime
from pathlib import Path

# Configuration
DROPBOX_URL = "https://www.dropbox.com/scl/fi/i0ji499omsoc0yl4fuzo7/STEPH-VERSION-ALL-ROOMS-2026.xlsx?rlkey=lq3j8vz7hdlnoqvairj0yp2ih&dl=1"
OUTPUT_FILE = "index.html"

# Rows Pia can book
ALLOWED_ROWS = [44, 48, 52, 56, 60, 140, 144, 148, 168, 184, 188, 192, 204, 208, 220, 224, 244, 248, 256, 264, 272, 284]

def download_excel():
    """Download Excel file from Dropbox."""
    print("📥 Downloading Excel from Dropbox...")
    response = requests.get(DROPBOX_URL)
    response.raise_for_status()
    
    temp_file = Path("temp_availability.xlsx")
    temp_file.write_bytes(response.content)
    print(f"✅ Downloaded {len(response.content)} bytes")
    return temp_file

def parse_availability(excel_path):
    """Parse Excel and extract room availability data."""
    print("📊 Parsing availability data...")
    df = pd.read_excel(excel_path, header=None)
    
    # Find all Saturdays (week starts)
    saturdays = []
    for col_idx in range(3, len(df.columns), 2):
        date_val = df.iloc[0, col_idx]
        day_name = df.iloc[1, col_idx]
        if pd.notna(day_name) and 'SATURDAY' in str(day_name).upper():
            if pd.notna(date_val):
                dt = pd.to_datetime(date_val)
                saturdays.append({
                    'col': col_idx,
                    'date': dt.strftime('%Y-%m-%d'),
                    'display': dt.strftime('%d %b'),
                    'weekNum': dt.isocalendar()[1]
                })
    
    # Extract room data
    rooms_data = []
    for row in ALLOWED_ROWS:
        building = str(df.iloc[row-1, 0]).strip() if pd.notna(df.iloc[row-1, 0]) else ""
        room_num = str(df.iloc[row-1, 1]).strip() if pd.notna(df.iloc[row-1, 1]) else ""
        
        # Clean up room name
        room_name = f"{building} {room_num}".strip()
        room_name = room_name.replace("     ", " ").replace("    ", " ").replace("   ", " ").replace("  ", " ")
        room_name = room_name.replace("EU OR UK  ROOM", "EU/UK").replace("EU ROOM", "EU")
        
        availability = {}
        for sat in saturdays:
            cell_value = df.iloc[row-1, sat['col']]
            is_available = pd.isna(cell_value) or str(cell_value).strip() == ""
            availability[sat['date']] = {
                'available': is_available,
                'bookedBy': "" if is_available else str(cell_value)[:25]
            }
        
        rooms_data.append({
            'row': row,
            'name': room_name,
            'building': building,
            'availability': availability
        })
    
    print(f"✅ Found {len(saturdays)} weeks and {len(rooms_data)} rooms")
    
    return {
        'generated': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'saturdays': saturdays,
        'rooms': rooms_data
    }

def generate_html(data):
    """Generate the interactive HTML dashboard."""
    print("🎨 Generating HTML dashboard...")
    
    data_json = json.dumps(data)
    
    html = f'''<!DOCTYPE html>
<html lang="sv">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Wildwind Rumstillgänglighet 2026</title>
    <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>⛵</text></svg>">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.5/babel.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #0c1929 0%, #1a3a5c 100%);
            min-height: 100vh;
            color: #e0e6ed;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
        header {{ 
            text-align: center; 
            padding: 30px 0;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            margin-bottom: 30px;
        }}
        h1 {{ 
            font-size: 2.5rem; 
            font-weight: 300;
            letter-spacing: 2px;
            margin-bottom: 10px;
            background: linear-gradient(90deg, #4facfe, #00f2fe);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}
        .subtitle {{ color: #7a8fa6; font-size: 0.9rem; }}
        .controls {{ 
            display: flex; 
            gap: 15px; 
            margin-bottom: 25px; 
            flex-wrap: wrap;
            justify-content: center;
        }}
        .filter-btn {{
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.15);
            color: #b8c5d4;
            padding: 10px 20px;
            border-radius: 25px;
            cursor: pointer;
            transition: all 0.3s;
            font-size: 0.9rem;
        }}
        .filter-btn:hover, .filter-btn.active {{
            background: linear-gradient(135deg, #4facfe, #00f2fe);
            border-color: transparent;
            color: #0c1929;
        }}
        .view-toggle {{
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            justify-content: center;
        }}
        .grid-view {{ display: grid; gap: 15px; }}
        .room-card {{
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 12px;
            padding: 20px;
            transition: all 0.3s;
        }}
        .room-card:hover {{
            background: rgba(255,255,255,0.06);
            border-color: rgba(79,172,254,0.3);
            transform: translateY(-2px);
        }}
        .room-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }}
        .room-name {{ 
            font-size: 1.1rem; 
            font-weight: 600;
            color: #fff;
        }}
        .building-tag {{
            background: rgba(79,172,254,0.15);
            color: #4facfe;
            padding: 4px 12px;
            border-radius: 15px;
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        .weeks-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(70px, 1fr));
            gap: 6px;
        }}
        .week-cell {{
            padding: 8px 4px;
            border-radius: 6px;
            text-align: center;
            font-size: 0.75rem;
            transition: all 0.2s;
            cursor: default;
        }}
        .week-cell.available {{
            background: linear-gradient(135deg, #00d26a, #00a854);
            color: #fff;
            font-weight: 500;
        }}
        .week-cell.booked {{
            background: rgba(255,255,255,0.05);
            color: #5a6a7a;
        }}
        .week-cell:hover {{
            transform: scale(1.05);
        }}
        .week-label {{ font-weight: 600; }}
        .week-num {{ font-size: 0.65rem; opacity: 0.7; }}
        
        .summary-section {{ margin-bottom: 40px; }}
        .summary-title {{
            font-size: 1.2rem;
            color: #4facfe;
            margin-bottom: 15px;
            padding-left: 10px;
            border-left: 3px solid #4facfe;
        }}
        .stats-bar {{
            display: flex;
            gap: 30px;
            justify-content: center;
            padding: 20px;
            background: rgba(255,255,255,0.03);
            border-radius: 12px;
            margin-bottom: 30px;
        }}
        .stat-item {{ text-align: center; }}
        .stat-value {{ 
            font-size: 2rem; 
            font-weight: 700;
            background: linear-gradient(90deg, #4facfe, #00f2fe);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}
        .stat-label {{ font-size: 0.8rem; color: #7a8fa6; }}
        
        .week-section {{ margin-bottom: 30px; }}
        .week-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px 20px;
            background: rgba(79,172,254,0.1);
            border-radius: 10px;
            margin-bottom: 10px;
        }}
        .week-date {{ font-size: 1.1rem; font-weight: 600; }}
        .week-count {{ 
            background: #00d26a;
            color: #fff;
            padding: 4px 12px;
            border-radius: 15px;
            font-size: 0.85rem;
        }}
        .rooms-list {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            padding-left: 20px;
        }}
        .room-tag {{
            background: rgba(255,255,255,0.08);
            padding: 8px 16px;
            border-radius: 8px;
            font-size: 0.85rem;
        }}
        .contact-btn {{
            display: inline-block;
            margin-top: 30px;
            padding: 15px 40px;
            background: linear-gradient(135deg, #4facfe, #00f2fe);
            color: #0c1929;
            text-decoration: none;
            border-radius: 30px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .contact-btn:hover {{
            transform: translateY(-2px);
            box-shadow: 0 10px 30px rgba(79,172,254,0.3);
        }}
        footer {{
            text-align: center;
            padding: 40px 0 20px;
            color: #5a6a7a;
            font-size: 0.85rem;
        }}
        
        @media (max-width: 768px) {{
            h1 {{ font-size: 1.8rem; }}
            .controls {{ flex-direction: column; align-items: stretch; }}
            .filter-btn {{ text-align: center; }}
            .stats-bar {{ flex-wrap: wrap; gap: 20px; }}
        }}
    </style>
</head>
<body>
    <div id="root"></div>
    
    <script type="text/babel">
        const data = {data_json};
        
        function App() {{
            const [view, setView] = React.useState('rooms');
            const [buildingFilter, setBuildingFilter] = React.useState('all');
            
            const buildings = [...new Set(data.rooms.map(r => r.building))];
            
            const filteredRooms = data.rooms.filter(room => {{
                if (buildingFilter !== 'all' && room.building !== buildingFilter) return false;
                return true;
            }});
            
            const totalSlots = data.rooms.length * data.saturdays.length;
            const availableSlots = data.rooms.reduce((sum, room) => 
                sum + data.saturdays.filter(sat => room.availability[sat.date]?.available).length, 0
            );
            
            const roomsPerWeek = data.saturdays.map(sat => ({{
                ...sat,
                rooms: data.rooms.filter(room => room.availability[sat.date]?.available)
            }}));
            
            return (
                <div className="container">
                    <header>
                        <h1>⛵ WILDWIND 2026</h1>
                        <p className="subtitle">Rumstillgänglighet • Uppdaterad {{data.generated}}</p>
                    </header>
                    
                    <div className="stats-bar">
                        <div className="stat-item">
                            <div className="stat-value">{{data.rooms.length}}</div>
                            <div className="stat-label">Rum tillgängliga</div>
                        </div>
                        <div className="stat-item">
                            <div className="stat-value">{{data.saturdays.length}}</div>
                            <div className="stat-label">Veckor</div>
                        </div>
                        <div className="stat-item">
                            <div className="stat-value">{{availableSlots}}</div>
                            <div className="stat-label">Lediga platser</div>
                        </div>
                        <div className="stat-item">
                            <div className="stat-value">{{Math.round(availableSlots/totalSlots*100)}}%</div>
                            <div className="stat-label">Tillgänglighet</div>
                        </div>
                    </div>
                    
                    <div className="view-toggle">
                        <button 
                            className={{`filter-btn ${{view === 'rooms' ? 'active' : ''}}`}}
                            onClick={{() => setView('rooms')}}
                        >
                            📋 Per rum
                        </button>
                        <button 
                            className={{`filter-btn ${{view === 'weeks' ? 'active' : ''}}`}}
                            onClick={{() => setView('weeks')}}
                        >
                            📅 Per vecka
                        </button>
                    </div>
                    
                    <div className="controls">
                        <button 
                            className={{`filter-btn ${{buildingFilter === 'all' ? 'active' : ''}}`}}
                            onClick={{() => setBuildingFilter('all')}}
                        >
                            Alla byggnader
                        </button>
                        {{buildings.map(b => (
                            <button 
                                key={{b}}
                                className={{`filter-btn ${{buildingFilter === b ? 'active' : ''}}`}}
                                onClick={{() => setBuildingFilter(b)}}
                            >
                                {{b}}
                            </button>
                        ))}}
                    </div>
                    
                    {{view === 'rooms' ? (
                        <div className="grid-view">
                            {{filteredRooms.map(room => (
                                <div key={{room.row}} className="room-card">
                                    <div className="room-header">
                                        <span className="room-name">{{room.name}}</span>
                                        <span className="building-tag">{{room.building}}</span>
                                    </div>
                                    <div className="weeks-grid">
                                        {{data.saturdays.map(sat => {{
                                            const avail = room.availability[sat.date];
                                            return (
                                                <div 
                                                    key={{sat.date}}
                                                    className={{`week-cell ${{avail?.available ? 'available' : 'booked'}}`}}
                                                    title={{avail?.available ? 'Ledig' : avail?.bookedBy}}
                                                >
                                                    <div className="week-label">{{sat.display}}</div>
                                                    <div className="week-num">v{{sat.weekNum}}</div>
                                                </div>
                                            );
                                        }})}}
                                    </div>
                                </div>
                            ))}}
                        </div>
                    ) : (
                        <div>
                            {{roomsPerWeek.map(week => (
                                <div key={{week.date}} className="week-section">
                                    <div className="week-header">
                                        <span className="week-date">
                                            {{week.display}} (v{{week.weekNum}})
                                        </span>
                                        <span className="week-count">
                                            {{week.rooms.length}} lediga rum
                                        </span>
                                    </div>
                                    <div className="rooms-list">
                                        {{week.rooms
                                            .filter(r => buildingFilter === 'all' || r.building === buildingFilter)
                                            .map(room => (
                                            <span key={{room.row}} className="room-tag">
                                                {{room.name}}
                                            </span>
                                        ))}}
                                    </div>
                                </div>
                            ))}}
                        </div>
                    )}}
                    
                    <footer>
                        <a href="mailto:pia@seafari.se?subject=Wildwind%20bokningsförfrågan" className="contact-btn">
                            ✉️ Skicka bokningsförfrågan
                        </a>
                        <p style={{{{marginTop: '30px'}}}}>
                            Sidan uppdateras automatiskt varje morgon kl 06:00<br/>
                            Vid frågor, kontakta <a href="mailto:pia@seafari.se" style={{{{color: '#4facfe'}}}}>pia@seafari.se</a>
                        </p>
                    </footer>
                </div>
            );
        }}
        
        ReactDOM.render(<App />, document.getElementById('root'));
    </script>
</body>
</html>'''
    
    return html

def main():
    print("🚀 Wildwind Availability Updater")
    print("=" * 40)
    
    try:
        # Download
        excel_file = download_excel()
        
        # Parse
        data = parse_availability(excel_file)
        
        # Generate HTML
        html = generate_html(data)
        
        # Write output
        output_path = Path(OUTPUT_FILE)
        output_path.write_text(html, encoding='utf-8')
        print(f"✅ Generated {OUTPUT_FILE}")
        
        # Cleanup
        Path("temp_availability.xlsx").unlink(missing_ok=True)
        
        print("=" * 40)
        print("🎉 Done!")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        raise

if __name__ == "__main__":
    main()
