import React, { useState, useEffect, useMemo } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, LineChart, Line, Legend 
} from 'recharts';
import { LayoutDashboard, TrendingUp, DollarSign, ShoppingCart, MousePointerClick, Activity, Play, ArrowLeft } from 'lucide-react';

function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [isExecuting, setIsExecuting] = useState(false);
  
  // States cho Drill-down UI
  const [selectedSKU, setSelectedSKU] = useState(null);
  const [selectedCampaign, setSelectedCampaign] = useState(null);

  const loadDashboardData = async () => {
    setLoading(true);
    try {
      const apiUrl = `http://${window.location.hostname}:8000`;
      const response = await fetch(`${apiUrl}/api/dashboard`);
      if (response.ok) {
        const jsonData = await response.json();
        setData(jsonData);
      } else {
        console.error('API Error:', await response.text());
      }
    } catch (error) {
      console.error('Error fetching dashboard data:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadDashboardData();
  }, []);

  const handleRunCode = async () => {
    setIsExecuting(true);
    try {
      const apiUrl = `http://${window.location.hostname}:8000`;
      const response = await fetch(`${apiUrl}/api/run-pipeline`, { method: 'POST' });
      if (!response.ok) {
        const errorData = await response.json();
        alert('Lỗi: ' + (errorData.detail || 'Không thể chạy pipeline'));
      } else {
        alert('Chạy Pipeline thành công! Giao diện sẽ tự động cập nhật dữ liệu mới.');
        await loadDashboardData();
      }
    } catch (error) {
      alert('Lỗi kết nối Backend: ' + error.message);
    } finally {
      setIsExecuting(false);
    }
  };

  const formatCurrency = (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val || 0);
  const formatPercent = (val) => `${((val || 0) * 100).toFixed(1)}%`;

  // Aggregations
  const skuStats = useMemo(() => {
    if (!data || !data.campaigns) return [];
    const stats = {};
    data.campaigns.forEach(c => {
      const sku = c.SKU || 'Unknown';
      if (!stats[sku]) {
        stats[sku] = { SKU: sku, Spend: 0, Sales: 0, Orders: 0, Clicks: 0, CampaignsCount: 0 };
      }
      stats[sku].Spend += (c.Spend || 0);
      stats[sku].Sales += (c.Sales || 0);
      stats[sku].Orders += (c.Orders || 0);
      stats[sku].Clicks += (c.Clicks || 0);
      stats[sku].CampaignsCount += 1;
    });
    return Object.values(stats).map(s => ({
      ...s,
      ACOS: s.Sales > 0 ? s.Spend / s.Sales : 0
    })).sort((a, b) => b.Spend - a.Spend); // Sort by Spend desc
  }, [data]);

  const getCampaignHistory = (campaignName) => {
    if (!data || !data.campaign_time_series || !data.date_blocks) return [];
    const tsData = data.campaign_time_series.find(c => c['Campaign Name'] === campaignName);
    if (!tsData) return [];
    
    return data.date_blocks.map(block => ({
      date: block,
      Spend: tsData[`TS_Spend_${block}`] || 0,
      Sales: tsData[`TS_Sales_${block}`] || 0
    }));
  };

  if (loading && !data) return <div className="loader">Đang tải dữ liệu Dashboard...</div>;
  if (!data) return <div className="loader">Không có dữ liệu. Vui lòng đảm bảo Backend (serve_webapp.py) đang chạy ở port 8000.</div>;

  const { summary, campaigns, source_file } = data;

  // --- RENDER BLOCKS ---

  const renderSKUList = () => (
    <div className="table-container">
      <table>
        <thead>
          <tr>
            <th>SKU</th>
            <th>Total Campaigns</th>
            <th>Total Spend</th>
            <th>Total Sales</th>
            <th>Avg ACOS</th>
            <th>Orders</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {skuStats.map((sku, i) => (
            <tr key={i} style={{ cursor: 'pointer' }} onClick={() => setSelectedSKU(sku.SKU)}>
              <td style={{ fontWeight: 'bold', color: '#60a5fa' }}>{sku.SKU}</td>
              <td>{sku.CampaignsCount}</td>
              <td>{formatCurrency(sku.Spend)}</td>
              <td style={{ color: 'var(--success)' }}>{formatCurrency(sku.Sales)}</td>
              <td style={{ color: sku.ACOS > 0.3 ? 'var(--danger)' : 'var(--success)' }}>
                {formatPercent(sku.ACOS)}
              </td>
              <td>{sku.Orders}</td>
              <td><button style={{ padding: '4px 8px', fontSize: '12px', borderRadius: '4px', background: '#3b82f6', color: 'white', border: 'none', cursor: 'pointer' }}>Xem Campaign</button></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );

  const renderCampaignList = () => {
    const filteredCampaigns = campaigns.filter(c => c.SKU === selectedSKU);
    return (
      <>
        <div style={{ marginBottom: '15px' }}>
          <button 
            onClick={() => setSelectedSKU(null)} 
            style={{ display: 'flex', alignItems: 'center', gap: '5px', background: 'transparent', border: '1px solid #475569', color: '#cbd5e1', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer' }}
          >
            <ArrowLeft size={16}/> Quay lại danh sách SKU
          </button>
        </div>
        <h3 style={{ marginBottom: '15px', color: '#60a5fa' }}>Các Campaign thuộc SKU: {selectedSKU}</h3>
        <div className="table-container">
          <table>
            <thead>
              <tr>
                <th>Campaign Name</th>
                <th>Status</th>
                <th>Spend</th>
                <th>Sales</th>
                <th>ACOS</th>
                <th>Health</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredCampaigns.map((c, i) => {
                let badgeClass = 'star';
                if (c['Health Tag'].includes('Bleeder')) badgeClass = 'bleeder';
                else if (c['Health Tag'].includes('Watch')) badgeClass = 'watch';
                else if (c['Health Tag'].includes('Dead') || c['Health Tag'].includes('Sleep')) badgeClass = 'dead';
                
                return (
                  <tr key={i} style={{ cursor: 'pointer' }} onClick={() => setSelectedCampaign(c['Campaign Name'])}>
                    <td style={{ color: '#e2e8f0' }}>{c['Campaign Name']}</td>
                    <td>{c.Status}</td>
                    <td>{formatCurrency(c.Spend)}</td>
                    <td style={{ color: 'var(--success)' }}>{formatCurrency(c.Sales)}</td>
                    <td style={{ color: c.ACOS > 0.3 ? 'var(--danger)' : 'var(--success)' }}>
                      {formatPercent(c.ACOS)}
                    </td>
                    <td><span className={`badge ${badgeClass}`}>{c['Health Tag']}</span></td>
                    <td><button style={{ padding: '4px 8px', fontSize: '12px', borderRadius: '4px', background: '#10b981', color: 'white', border: 'none', cursor: 'pointer' }}>Xem Lịch sử</button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </>
    );
  };

  const renderCampaignHistory = () => {
    const historyData = getCampaignHistory(selectedCampaign);
    return (
      <>
        <div style={{ marginBottom: '15px' }}>
          <button 
            onClick={() => setSelectedCampaign(null)} 
            style={{ display: 'flex', alignItems: 'center', gap: '5px', background: 'transparent', border: '1px solid #475569', color: '#cbd5e1', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer' }}
          >
            <ArrowLeft size={16}/> Quay lại danh sách Campaign
          </button>
        </div>
        <h3 style={{ marginBottom: '15px', color: '#10b981' }}>Chi tiết Lịch sử Campaign: {selectedCampaign}</h3>
        
        {/* Line Chart */}
        {historyData.length > 0 ? (
          <div style={{ height: '300px', marginBottom: '30px' }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={historyData} margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                <XAxis dataKey="date" tick={{ fill: '#94a3b8' }} />
                <YAxis yAxisId="left" tick={{ fill: '#94a3b8' }} />
                <YAxis yAxisId="right" orientation="right" tick={{ fill: '#94a3b8' }} />
                <Tooltip contentStyle={{ backgroundColor: '#1e293b', border: 'none', borderRadius: '8px' }} itemStyle={{ color: '#f8fafc' }} />
                <Legend />
                <Line yAxisId="left" type="monotone" dataKey="Sales" stroke="#10b981" activeDot={{ r: 8 }} name="Sales ($)" />
                <Line yAxisId="right" type="monotone" dataKey="Spend" stroke="#ef4444" name="Spend ($)" />
              </LineChart>
            </ResponsiveContainer>
          </div>
        ) : (
          <p style={{ color: '#94a3b8', fontStyle: 'italic', marginBottom: '20px' }}>Không có dữ liệu thời gian cho Campaign này.</p>
        )}

        {/* History Table */}
        <div className="table-container">
          <table>
            <thead>
              <tr>
                <th>Khoảng thời gian (Date Block)</th>
                <th>Spend</th>
                <th>Sales</th>
                <th>ACOS</th>
              </tr>
            </thead>
            <tbody>
              {historyData.map((d, i) => (
                <tr key={i}>
                  <td style={{ fontWeight: 'bold' }}>{d.date}</td>
                  <td style={{ color: '#ef4444' }}>{formatCurrency(d.Spend)}</td>
                  <td style={{ color: '#10b981' }}>{formatCurrency(d.Sales)}</td>
                  <td style={{ color: (d.Sales > 0 ? (d.Spend/d.Sales) : 0) > 0.3 ? 'var(--danger)' : 'var(--success)' }}>
                    {d.Sales > 0 ? formatPercent(d.Spend / d.Sales) : '0.0%'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </>
    );
  };

  return (
    <div className="app-container">
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <h1><LayoutDashboard size={28} style={{verticalAlign: 'bottom', marginRight: '8px'}}/> Amazon Ads Analytics</h1>
          <p>Nguồn dữ liệu: {source_file}</p>
        </div>
        <button 
          onClick={handleRunCode} 
          disabled={isExecuting}
          style={{ 
            padding: '10px 20px', 
            borderRadius: '8px', 
            cursor: isExecuting ? 'not-allowed' : 'pointer', 
            display: 'flex', 
            alignItems: 'center', 
            gap: '8px', 
            fontWeight: 'bold',
            backgroundColor: isExecuting ? '#64748b' : '#3b82f6',
            color: 'white',
            border: 'none',
            opacity: isExecuting ? 0.7 : 1
          }}
        >
          <Play size={18} />
          {isExecuting ? 'Đang chạy Pipeline...' : 'Run Full Code'}
        </button>
      </header>

      {/* Summary Cards */}
      <section className="summary-grid">
        <div className="glass-panel stat-card">
          <div className="stat-title"><DollarSign size={16} /> Total Spend</div>
          <div className="stat-value">{formatCurrency(summary.total_spend)}</div>
        </div>
        <div className="glass-panel stat-card">
          <div className="stat-title"><TrendingUp size={16} /> Total Sales</div>
          <div className="stat-value success">{formatCurrency(summary.total_sales)}</div>
        </div>
        <div className="glass-panel stat-card">
          <div className="stat-title"><Activity size={16} /> Overall ACOS</div>
          <div className={`stat-value ${summary.overall_acos > 0.3 ? 'danger' : 'success'}`}>
            {formatPercent(summary.overall_acos)}
          </div>
        </div>
        <div className="glass-panel stat-card">
          <div className="stat-title"><ShoppingCart size={16} /> Orders</div>
          <div className="stat-value">{summary.total_orders}</div>
        </div>
        <div className="glass-panel stat-card">
          <div className="stat-title"><MousePointerClick size={16} /> Clicks</div>
          <div className="stat-value">{summary.total_clicks}</div>
        </div>
      </section>

      {/* Main Content Area */}
      <section className="glass-panel">
        <h2 style={{ marginBottom: '20px', fontSize: '18px', color: 'var(--text-secondary)' }}>
          {!selectedSKU && !selectedCampaign && "Drill-down: Danh sách SKU"}
          {selectedSKU && !selectedCampaign && "Drill-down: Danh sách Campaign"}
          {selectedCampaign && "Drill-down: Lịch sử Campaign"}
        </h2>
        
        {/* Routing Mechanism */}
        {!selectedSKU && !selectedCampaign && renderSKUList()}
        {selectedSKU && !selectedCampaign && renderCampaignList()}
        {selectedCampaign && renderCampaignHistory()}
      </section>

    </div>
  );
}

export default App;
