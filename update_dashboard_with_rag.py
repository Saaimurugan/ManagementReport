import json
import re

# Read the existing standalone dashboard
with open('dashboard_standalone.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

# Add RAG status styles after the existing styles
rag_styles = """
        /* RAG Status Section */
        .rag-container {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            margin-bottom: 30px;
        }

        .rag-title {
            font-size: 1.8em;
            font-weight: 700;
            color: #333;
            margin-bottom: 25px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .rag-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 25px;
        }

        .rag-card {
            background: linear-gradient(135deg, #f5f7fa 0%, #ffffff 100%);
            padding: 25px;
            border-radius: 12px;
            border: 2px solid #e0e0e0;
            transition: all 0.3s;
            position: relative;
            overflow: hidden;
        }

        .rag-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 5px;
            height: 100%;
            background: var(--status-color);
        }

        .rag-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        }

        .rag-card.status-green { --status-color: #00b894; }
        .rag-card.status-amber { --status-color: #fdcb6e; }
        .rag-card.status-red { --status-color: #d63031; }

        .rag-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }

        .rag-metric-name {
            font-size: 1.3em;
            font-weight: 600;
            color: #333;
        }

        .rag-indicator {
            width: 50px;
            height: 50px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 1.2em;
            color: white;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        }

        .rag-indicator.green { background: linear-gradient(135deg, #00b894, #55efc4); }
        .rag-indicator.amber { background: linear-gradient(135deg, #fdcb6e, #ffeaa7); }
        .rag-indicator.red { background: linear-gradient(135deg, #d63031, #ff7675); }

        .rag-details {
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #e0e0e0;
        }

        .rag-detail-row {
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
            font-size: 0.95em;
        }

        .rag-detail-label {
            color: #666;
        }

        .rag-detail-value {
            font-weight: 600;
            color: #333;
        }

        .rag-status-message {
            margin-top: 15px;
            padding: 12px;
            background: rgba(0,0,0,0.03);
            border-radius: 8px;
            font-size: 0.9em;
            color: #555;
            line-height: 1.5;
        }

        .rag-legend {
            display: flex;
            gap: 30px;
            margin-top: 20px;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 10px;
            flex-wrap: wrap;
        }

        .rag-legend-item {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.9em;
        }

        .rag-legend-color {
            width: 20px;
            height: 20px;
            border-radius: 50%;
        }

        .rag-legend-color.green { background: #00b894; }
        .rag-legend-color.amber { background: #fdcb6e; }
        .rag-legend-color.red { background: #d63031; }
"""

# Find the closing </style> tag and insert before it
html_content = html_content.replace('</style>', rag_styles + '\n    </style>')

# Add RAG HTML section after filters, before metrics
rag_html = """
        <div class="rag-container">
            <div class="rag-title">
                🎯 Project Health Status
            </div>
            
            <div class="rag-grid" id="ragGrid">
                <!-- RAG cards will be inserted here -->
            </div>

            <div class="rag-legend">
                <div class="rag-legend-item">
                    <div class="rag-legend-color green"></div>
                    <strong>Green:</strong> On track - No issues
                </div>
                <div class="rag-legend-item">
                    <div class="rag-legend-color amber"></div>
                    <strong>Amber:</strong> At risk - Needs attention
                </div>
                <div class="rag-legend-item">
                    <div class="rag-legend-color red"></div>
                    <strong>Red:</strong> Critical - Immediate action required
                </div>
            </div>
        </div>
"""

# Insert RAG HTML after the filters div
html_content = html_content.replace(
    '<div class="metrics" id="metrics"></div>',
    rag_html + '\n        <div class="metrics" id="metrics"></div>'
)

# Add RAG calculation function before initializeDashboard
rag_js = """
        function calculateRAGStatus() {
            const ragGrid = document.getElementById('ragGrid');
            
            // Calculate Time Status
            const totalTasks = filteredData.length;
            const completedTasks = filteredData.filter(item => item.Status === 'Done').length;
            const inProgressTasks = filteredData.filter(item => item.Status === 'In Progress').length;
            const completionRate = totalTasks > 0 ? (completedTasks / totalTasks) * 100 : 0;
            
            let timeStatus = 'green';
            let timeMessage = 'Project timeline is on track with strong progress.';
            if (completionRate < 30) {
                timeStatus = 'red';
                timeMessage = 'Critical delay: Only ' + completionRate.toFixed(1) + '% complete. Immediate action needed.';
            } else if (completionRate < 60) {
                timeStatus = 'amber';
                timeMessage = 'Timeline at risk: ' + completionRate.toFixed(1) + '% complete. Monitor closely.';
            }

            // Calculate Cost Status
            const totalBudgetPlanned = filteredData.reduce((sum, item) => sum + item['Budget Planned'], 0);
            const totalBudgetConsumed = filteredData.reduce((sum, item) => sum + item['Budget Consumed'], 0);
            const totalBudgetRemaining = filteredData.reduce((sum, item) => sum + item['Budget Remaining'], 0);
            const budgetUtilization = totalBudgetPlanned > 0 ? (totalBudgetConsumed / totalBudgetPlanned) * 100 : 0;
            const budgetVariance = totalBudgetRemaining >= 0 ? 'Under budget' : 'Over budget';
            
            let costStatus = 'green';
            let costMessage = 'Budget is well managed and under control.';
            if (totalBudgetRemaining < 0 || budgetUtilization > 95) {
                costStatus = 'red';
                costMessage = 'Critical: Budget overspent by $' + Math.abs(totalBudgetRemaining).toLocaleString() + '. Immediate review required.';
            } else if (budgetUtilization > 80 || totalBudgetRemaining < totalBudgetPlanned * 0.1) {
                costStatus = 'amber';
                costMessage = 'Warning: High budget utilization at ' + budgetUtilization.toFixed(1) + '%. Monitor spending.';
            }

            // Calculate Quality Status
            const highQualityCount = filteredData.filter(item => item.Quality === 'High').length;
            const lowQualityCount = filteredData.filter(item => item.Quality === 'Low').length;
            const qualityScore = totalTasks > 0 ? ((highQualityCount / totalTasks) * 100) : 0;
            const lowQualityPercent = totalTasks > 0 ? ((lowQualityCount / totalTasks) * 100) : 0;
            
            let qualityStatus = 'green';
            let qualityMessage = 'Quality standards are being met consistently.';
            if (lowQualityPercent > 40 || qualityScore < 20) {
                qualityStatus = 'red';
                qualityMessage = 'Critical quality issues: ' + lowQualityCount + ' low quality items (' + lowQualityPercent.toFixed(1) + '%). Immediate review needed.';
            } else if (lowQualityPercent > 25 || qualityScore < 35) {
                qualityStatus = 'amber';
                qualityMessage = 'Quality concerns: Monitor the ' + lowQualityCount + ' low quality items closely.';
            }

            // Generate RAG cards HTML
            ragGrid.innerHTML = `
                <div class="rag-card status-${timeStatus}">
                    <div class="rag-header">
                        <div class="rag-metric-name">⏱️ Time</div>
                        <div class="rag-indicator ${timeStatus}">${timeStatus === 'green' ? '✓' : timeStatus === 'amber' ? '!' : '✗'}</div>
                    </div>
                    <div class="rag-details">
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Completion Rate:</span>
                            <span class="rag-detail-value">${completionRate.toFixed(1)}%</span>
                        </div>
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Completed Tasks:</span>
                            <span class="rag-detail-value">${completedTasks} / ${totalTasks}</span>
                        </div>
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">In Progress:</span>
                            <span class="rag-detail-value">${inProgressTasks}</span>
                        </div>
                    </div>
                    <div class="rag-status-message">${timeMessage}</div>
                </div>

                <div class="rag-card status-${costStatus}">
                    <div class="rag-header">
                        <div class="rag-metric-name">💰 Cost</div>
                        <div class="rag-indicator ${costStatus}">${costStatus === 'green' ? '✓' : costStatus === 'amber' ? '!' : '✗'}</div>
                    </div>
                    <div class="rag-details">
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Budget Utilization:</span>
                            <span class="rag-detail-value">${budgetUtilization.toFixed(1)}%</span>
                        </div>
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Consumed:</span>
                            <span class="rag-detail-value">$${totalBudgetConsumed.toLocaleString()}</span>
                        </div>
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Remaining:</span>
                            <span class="rag-detail-value ${totalBudgetRemaining >= 0 ? 'positive' : 'negative'}">$${totalBudgetRemaining.toLocaleString()}</span>
                        </div>
                    </div>
                    <div class="rag-status-message">${costMessage}</div>
                </div>

                <div class="rag-card status-${qualityStatus}">
                    <div class="rag-header">
                        <div class="rag-metric-name">⭐ Quality</div>
                        <div class="rag-indicator ${qualityStatus}">${qualityStatus === 'green' ? '✓' : qualityStatus === 'amber' ? '!' : '✗'}</div>
                    </div>
                    <div class="rag-details">
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">High Quality:</span>
                            <span class="rag-detail-value">${highQualityCount} (${qualityScore.toFixed(1)}%)</span>
                        </div>
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Medium Quality:</span>
                            <span class="rag-detail-value">${filteredData.filter(item => item.Quality === 'Medium').length}</span>
                        </div>
                        <div class="rag-detail-row">
                            <span class="rag-detail-label">Low Quality:</span>
                            <span class="rag-detail-value">${lowQualityCount} (${lowQualityPercent.toFixed(1)}%)</span>
                        </div>
                    </div>
                    <div class="rag-status-message">${qualityMessage}</div>
                </div>
            `;
        }
"""

# Insert RAG function before initializeDashboard
html_content = html_content.replace(
    'function initializeDashboard() {',
    rag_js + '\n        function initializeDashboard() {'
)

# Update the updateDashboard function to call RAG calculation
html_content = html_content.replace(
    'function updateDashboard() {\n            updateMetrics();',
    'function updateDashboard() {\n            calculateRAGStatus();\n            updateMetrics();'
)

# Write the updated HTML
with open('dashboard_standalone.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("✅ Successfully updated dashboard with RAG status indicators!")
print("📊 Added Time, Cost, and Quality health metrics with Green/Amber/Red indicators")
