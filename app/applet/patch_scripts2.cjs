const fs = require('fs');
let code = fs.readFileSync('gas-project/Scripts.html', 'utf8');

// 1. Add loadPlanning to navigation
const navTarget = "if(pageId === 'page-transactions') loadTransactions();";
const navReplace = "if(pageId === 'page-transactions') loadTransactions();\n    if(pageId === 'page-planning') loadPlanning();";
code = code.replace(navTarget, navReplace);

// 2. Append Planning logic at the end
const planningJS = `
  // ==========================================
  // PLANEJAMENTO FINANCEIRO ANUAL
  // ==========================================
  let currentPlanningYear = new Date().getFullYear();
  let planningEvolutionChart = null;
  let planningCompositionChart = null;
  let planningDataCache = null;

  function initPlanningYearSelect() {
    const select = $('#planning-year-select');
    if (select.children().length > 0) return;
    const current = new Date().getFullYear();
    for (let y = current - 2; y <= current + 2; y++) {
      select.append(\`<option value="\${y}" \${y === current ? 'selected' : ''}>\${y}</option>\`);
    }
  }

  function loadPlanning() {
    initPlanningYearSelect();
    $('#planning-loading').removeClass('hidden');
    $('#planning-matrix-container').addClass('hidden');
    
    currentPlanningYear = parseInt($('#planning-year-select').val());

    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run
        .withSuccessHandler(function(data) {
          planningDataCache = data;
          renderPlanning(data);
          $('#planning-loading').addClass('hidden');
          $('#planning-matrix-container').removeClass('hidden');
        })
        .withFailureHandler(function(err) {
          console.error(err);
          alert('Erro ao carregar planejamento.');
          $('#planning-loading').addClass('hidden');
        })
        .getPlanningData(currentPlanningYear);
    } else {
      setTimeout(() => {
        const mockData = generateMockPlanningData(currentPlanningYear);
        planningDataCache = mockData;
        renderPlanning(mockData);
        $('#planning-loading').addClass('hidden');
        $('#planning-matrix-container').removeClass('hidden');
      }, 500);
    }
  }

  $('#planning-year-select').change(function() {
    loadPlanning();
  });
  $('#btn-refresh-planning').click(function() {
    loadPlanning();
  });

  function formatCell(valObj) {
    if (!valObj || valObj.total === 0) return '-';
    return valObj.total.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function renderPlanning(data) {
    const months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    if (data.year === new Date().getFullYear()) {
      $('#planning-current-month-badge').removeClass('hidden');
      $('#planning-current-month-text').text(\`\${months[data.currentMonthNum - 1]}/\${data.year}\`);
    } else {
      $('#planning-current-month-badge').addClass('hidden');
    }

    const sum = data.summary;
    $('#plan-summary-income').text(formatCurrency(sum.incomeRealized + sum.incomePlanned));
    $('#plan-summary-income-realized').text(formatCurrency(sum.incomeRealized));
    $('#plan-summary-income-planned').text(formatCurrency(sum.incomePlanned));
    
    $('#plan-summary-expense').text(formatCurrency(sum.expenseRealized + sum.expensePlanned));
    $('#plan-summary-expense-realized').text(formatCurrency(sum.expenseRealized));
    $('#plan-summary-expense-planned').text(formatCurrency(sum.expensePlanned));
    
    const currBalEl = $('#plan-summary-current-balance');
    currBalEl.text(formatCurrency(sum.currentBalance));
    currBalEl.removeClass('text-gray-900 dark:text-white text-rose-600 dark:text-rose-400');
    if (sum.currentBalance < 0) currBalEl.addClass('text-rose-600 dark:text-rose-400');
    else currBalEl.addClass('text-gray-900 dark:text-white');

    const projBalEl = $('#plan-summary-projected-balance');
    projBalEl.text(formatCurrency(sum.projectedBalance));
    projBalEl.removeClass('text-gray-900 dark:text-white text-rose-600 dark:text-rose-400 text-purple-600 dark:text-purple-400');
    if (sum.projectedBalance < 0) projBalEl.addClass('text-rose-600 dark:text-rose-400');
    else if (sum.projectedBalance > 0) projBalEl.addClass('text-purple-600 dark:text-purple-400');
    else projBalEl.addClass('text-gray-900 dark:text-white');

    let theadHtml = \`<tr><th class="px-4 py-3 sticky left-0 z-10 bg-gray-50 dark:bg-gray-900 border-r border-gray-200 dark:border-gray-700 min-w-[200px]">Categoria</th>\`;
    data.monthsMeta.forEach(m => {
      let style = '';
      if (m.status === 'realized') style = 'text-gray-400 dark:text-gray-500 font-normal';
      else if (m.status === 'current') style = 'bg-blue-100 dark:bg-blue-900/30 text-blue-800 dark:text-blue-300 ring-2 ring-inset ring-primary font-bold rounded-t-md';
      else style = 'text-purple-600 dark:text-purple-400 font-medium';
      
      theadHtml += \`<th class="px-3 py-3 text-center min-w-[100px] \${style}">\${m.name}</th>\`;
    });
    theadHtml += \`<th class="px-4 py-3 text-right bg-gray-50 dark:bg-gray-900 font-bold border-l border-gray-200 dark:border-gray-700">Total Ano</th></tr>\`;
    $('#planning-thead').html(theadHtml);

    let tbodyHtml = '';

    // RECEITAS
    let recMasterHtml = \`<tr class="planning-toggle-row cursor-pointer select-none transition-colors bg-emerald-50 dark:bg-emerald-900/20 text-emerald-800 dark:text-emerald-300" data-target="row-receitas">
      <td class="px-4 py-3 sticky left-0 z-10 font-bold flex items-center justify-between bg-emerald-50 dark:bg-emerald-900/20 border-r border-gray-200 dark:border-gray-700">
        <span>Receitas</span>
        <i class="fa-solid fa-chevron-down text-xs transition-transform duration-200"></i>
      </td>\`;
    
    let recTotals = Array(12).fill(0);
    let recYearTotal = 0;
    let recChildren = '';
    
    Object.keys(data.incomeCategories).sort().forEach(cat => {
      let catYearTotal = 0;
      let cellsHtml = '';
      data.incomeCategories[cat].forEach((m, idx) => {
        recTotals[idx] += m.total;
        catYearTotal += m.total;
        cellsHtml += \`<td class="px-3 py-2 text-center planning-cell-action cursor-pointer hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors" data-cat="\${cat}" data-month="\${idx+1}" data-tipo="Entrada">\${formatCell(m)}</td>\`;
      });
      recYearTotal += catYearTotal;
      recChildren += \`<tr class="row-receitas hidden bg-white dark:bg-gray-800 hover:bg-gray-50 dark:hover:bg-gray-750 transition-colors group">
        <td class="px-4 py-2 sticky left-0 z-10 bg-white dark:bg-gray-800 group-hover:bg-gray-50 dark:group-hover:bg-gray-750 border-r border-gray-200 dark:border-gray-700 text-gray-700 dark:text-gray-300 pl-8">\${cat}</td>
        \${cellsHtml}
        <td class="px-4 py-2 text-right font-semibold text-emerald-600 dark:text-emerald-400 border-l border-gray-200 dark:border-gray-700 bg-gray-50/50 dark:bg-gray-900/30">\${formatCurrency(catYearTotal)}</td>
      </tr>\`;
    });

    recTotals.forEach(t => {
      recMasterHtml += \`<td class="px-3 py-3 text-center font-bold">\${t === 0 ? '-' : t.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>\`;
    });
    recMasterHtml += \`<td class="px-4 py-3 text-right font-bold border-l border-gray-200 dark:border-gray-700">\${formatCurrency(recYearTotal)}</td></tr>\`;
    
    tbodyHtml += recMasterHtml + recChildren;

    // DESPESAS
    const renderMacro = (macroName, rowClass, colorClasses, colorClassesTd, textClassTotal) => {
      if (!data.expenseMacros[macroName] || Object.keys(data.expenseMacros[macroName]).length === 0) return;
      
      let masterHtml = \`<tr class="planning-toggle-row cursor-pointer select-none transition-colors \${colorClasses}" data-target="\${rowClass}">
        <td class="px-4 py-3 sticky left-0 z-10 font-bold flex items-center justify-between \${colorClassesTd} border-r border-gray-200 dark:border-gray-700">
          <span>\${macroName}</span>
          <i class="fa-solid fa-chevron-down text-xs transition-transform duration-200"></i>
        </td>\`;
      
      let totals = Array(12).fill(0);
      let yearTotal = 0;
      let childrenHtml = '';

      Object.keys(data.expenseMacros[macroName]).sort().forEach(cat => {
        let catYearTotal = 0;
        let cellsHtml = '';
        data.expenseMacros[macroName][cat].forEach((m, idx) => {
          totals[idx] += m.total;
          catYearTotal += m.total;
          cellsHtml += \`<td class="px-3 py-2 text-center planning-cell-action cursor-pointer hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors" data-cat="\${cat}" data-month="\${idx+1}" data-tipo="Saída">\${formatCell(m)}</td>\`;
        });
        yearTotal += catYearTotal;
        childrenHtml += \`<tr class="\${rowClass} hidden bg-white dark:bg-gray-800 hover:bg-gray-50 dark:hover:bg-gray-750 transition-colors group">
          <td class="px-4 py-2 sticky left-0 z-10 bg-white dark:bg-gray-800 group-hover:bg-gray-50 dark:group-hover:bg-gray-750 border-r border-gray-200 dark:border-gray-700 text-gray-700 dark:text-gray-300 pl-8">\${cat}</td>
          \${cellsHtml}
          <td class="px-4 py-2 text-right font-semibold \${textClassTotal} border-l border-gray-200 dark:border-gray-700 bg-gray-50/50 dark:bg-gray-900/30">\${formatCurrency(catYearTotal)}</td>
        </tr>\`;
      });

      totals.forEach(t => {
        masterHtml += \`<td class="px-3 py-3 text-center font-bold">\${t === 0 ? '-' : t.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>\`;
      });
      masterHtml += \`<td class="px-4 py-3 text-right font-bold border-l border-gray-200 dark:border-gray-700">\${formatCurrency(yearTotal)}</td></tr>\`;
      
      tbodyHtml += masterHtml + childrenHtml;
    };

    renderMacro('Fixas', 'row-fixas', 'bg-rose-50 dark:bg-rose-900/20 text-rose-800 dark:text-rose-300', 'bg-rose-50 dark:bg-rose-900/20', 'text-rose-600 dark:text-rose-400');
    renderMacro('Variáveis Essenciais', 'row-essenciais', 'bg-orange-50 dark:bg-orange-900/20 text-orange-800 dark:text-orange-300', 'bg-orange-50 dark:bg-orange-900/20', 'text-orange-600 dark:text-orange-400');
    renderMacro('Discricionárias', 'row-discricionarias', 'bg-amber-50 dark:bg-amber-900/20 text-amber-800 dark:text-amber-300', 'bg-amber-50 dark:bg-amber-900/20', 'text-amber-600 dark:text-amber-400');

    $('#planning-tbody').html(tbodyHtml);

    let tfootHtml = \`<tr>
      <td class="px-4 py-3 sticky left-0 z-10 bg-gray-100 dark:bg-gray-900 border-r border-gray-300 dark:border-gray-600">Saldo do Mês</td>\`;
    let totalBalanceYear = 0;
    data.monthlyTotals.forEach(m => {
      totalBalanceYear += m.balance;
      let colorClass = m.balance < 0 ? 'text-rose-600 dark:text-rose-400' : 'text-gray-900 dark:text-white';
      tfootHtml += \`<td class="px-3 py-3 text-center \${colorClass}">\${m.balance === 0 ? '-' : m.balance.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>\`;
    });
    let totalColorClass = totalBalanceYear < 0 ? 'text-rose-600 dark:text-rose-400' : 'text-gray-900 dark:text-white';
    tfootHtml += \`<td class="px-4 py-3 text-right \${totalColorClass} border-l border-gray-300 dark:border-gray-600">\${formatCurrency(totalBalanceYear)}</td></tr>\`;
    
    $('#planning-tfoot').html(tfootHtml);

    renderPlanningCharts(data);
  }

  $(document).on('click', '.planning-toggle-row', function() {
    const targetClass = $(this).data('target');
    const icon = $(this).find('i.fa-chevron-down');
    
    $(\`.\${targetClass}\`).toggleClass('hidden');
    
    if ($(\`.\${targetClass}\`).first().hasClass('hidden')) {
      icon.css('transform', 'rotate(0deg)');
    } else {
      icon.css('transform', 'rotate(-180deg)');
    }
  });

  $(document).on('click', '.planning-cell-action', function() {
    const cat = $(this).data('cat');
    const month = parseInt($(this).data('month'));
    const tipo = $(this).data('tipo');
    const year = currentPlanningYear;

    $('#planning-drawer-title').text(cat);
    $('#planning-drawer-subtitle').text(\`Mês \${month}/\${year}\`);
    
    $('#planning-drawer-tx-list').empty();
    $('#planning-drawer-empty').addClass('hidden');
    $('#planning-drawer-loading').removeClass('hidden');

    $('#modal-planning-drawer').removeClass('hidden');
    
    setTimeout(() => {
      $('#planning-drawer-backdrop').removeClass('opacity-0');
      $('#planning-drawer-panel').removeClass('translate-x-full');
    }, 10);

    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run
        .withSuccessHandler(function(rows) {
          $('#planning-drawer-loading').addClass('hidden');
          renderDrawerTransactions(rows, cat, month, year, tipo);
        })
        .withFailureHandler(function(err) {
          console.error(err);
          $('#planning-drawer-loading').addClass('hidden');
          $('#planning-drawer-empty').removeClass('hidden');
        })
        .getTransactions(year.toString(), month.toString());
    } else {
      setTimeout(() => {
        $('#planning-drawer-loading').addClass('hidden');
        renderDrawerTransactions([], cat, month, year, tipo);
      }, 500);
    }
  });

  function renderDrawerTransactions(rows, cat, month, year, tipo) {
    const list = $('#planning-drawer-tx-list');
    list.empty();
    
    let filtered = rows.filter(r => r.categoria === cat && r.tipo === tipo);

    filtered = filtered.filter(r => {
      const d = new Date(r.data);
      return d.getFullYear() === year && (d.getMonth() + 1) === month;
    });

    if (filtered.length === 0) {
      $('#planning-drawer-empty').removeClass('hidden');
    } else {
      $('#planning-drawer-empty').addClass('hidden');
      filtered.forEach(t => {
        const d = new Date(t.data);
        const dateStr = \`\${d.getDate().toString().padStart(2, '0')}/\${(d.getMonth()+1).toString().padStart(2, '0')}\`;
        const color = t.tipo === 'Entrada' ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-400';
        list.append(\`
          <li class="py-3 flex justify-between items-center group">
            <div>
              <p class="text-sm font-medium text-gray-900 dark:text-white">\${t.descricao}</p>
              <p class="text-xs text-gray-500 dark:text-gray-400 mt-0.5">\${dateStr} &bull; \${t.origem || 'Manual'}</p>
            </div>
            <div class="text-sm font-bold \${color}">
              \${formatCurrency(t.valor)}
            </div>
          </li>
        \`);
      });
    }
  }

  $('#btn-close-planning-drawer, #planning-drawer-backdrop').click(function() {
    $('#planning-drawer-backdrop').addClass('opacity-0');
    $('#planning-drawer-panel').addClass('translate-x-full');
    setTimeout(() => {
      $('#modal-planning-drawer').addClass('hidden');
    }, 300);
  });

  function renderPlanningCharts(data) {
    if (planningEvolutionChart) planningEvolutionChart.destroy();
    if (planningCompositionChart) planningCompositionChart.destroy();

    const ctxEvol = document.getElementById('planningEvolutionChart').getContext('2d');
    const labels = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    const incomes = data.monthlyTotals.map(m => m.income);
    const expenses = data.monthlyTotals.map(m => m.expense);
    const accumBalance = [];
    let currentAccum = 0;
    data.monthlyTotals.forEach(m => {
      currentAccum += m.balance;
      accumBalance.push(currentAccum);
    });

    const isDark = document.documentElement.classList.contains('dark');
    const textColor = isDark ? '#94a3b8' : '#64748b';
    const gridColor = isDark ? '#334155' : '#e2e8f0';

    planningEvolutionChart = new Chart(ctxEvol, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Receitas',
            data: incomes,
            backgroundColor: '#10b981',
            borderRadius: 4,
            order: 2
          },
          {
            label: 'Despesas',
            data: expenses,
            backgroundColor: '#f43f5e',
            borderRadius: 4,
            order: 3
          },
          {
            label: 'Saldo Acumulado',
            data: accumBalance,
            type: 'line',
            borderColor: '#3b82f6',
            backgroundColor: '#3b82f6',
            borderWidth: 2,
            tension: 0.3,
            pointRadius: 3,
            order: 1
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: textColor } },
          tooltip: {
            callbacks: {
              label: function(context) {
                return context.dataset.label + ': ' + formatCurrency(context.raw);
              }
            }
          }
        },
        scales: {
          y: { ticks: { color: textColor }, grid: { color: gridColor } },
          x: { ticks: { color: textColor }, grid: { display: false } }
        }
      }
    });

    const ctxComp = document.getElementById('planningCompositionChart').getContext('2d');
    let sumFixas = 0;
    let sumEssenciais = 0;
    let sumDiscricionarias = 0;

    if (data.expenseMacros['Fixas']) {
      Object.values(data.expenseMacros['Fixas']).forEach(cat => cat.forEach(m => sumFixas += m.total));
    }
    if (data.expenseMacros['Variáveis Essenciais']) {
      Object.values(data.expenseMacros['Variáveis Essenciais']).forEach(cat => cat.forEach(m => sumEssenciais += m.total));
    }
    if (data.expenseMacros['Discricionárias']) {
      Object.values(data.expenseMacros['Discricionárias']).forEach(cat => cat.forEach(m => sumDiscricionarias += m.total));
    }

    planningCompositionChart = new Chart(ctxComp, {
      type: 'doughnut',
      data: {
        labels: ['Fixas', 'Variáveis Essenciais', 'Discricionárias'],
        datasets: [{
          data: [sumFixas, sumEssenciais, sumDiscricionarias],
          backgroundColor: ['#f43f5e', '#f97316', '#f59e0b'],
          borderWidth: 2,
          borderColor: isDark ? '#1f2937' : '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '65%',
        plugins: {
          legend: { position: 'right', labels: { color: textColor, padding: 20 } },
          tooltip: {
            callbacks: {
              label: function(context) {
                let sum = context.dataset.data.reduce((a, b) => a + b, 0);
                let pct = sum > 0 ? Math.round((context.raw / sum) * 100) + '%' : '0%';
                return \` \${context.label}: \${formatCurrency(context.raw)} (\${pct})\`;
              }
            }
          }
        }
      }
    });
  }

  function generateMockPlanningData(year) {
    return {
      year: year,
      currentMonthNum: new Date().getMonth() + 1,
      monthsMeta: Array.from({length:12}, (_, i) => ({month: i+1, status: i < new Date().getMonth() ? 'realized' : i === new Date().getMonth() ? 'current' : 'planned', name: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'][i]})),
      incomeCategories: {
        'Salário': Array.from({length:12}, () => ({realized:5000, planned:0, total:5000}))
      },
      expenseMacros: { 
        'Fixas': { 'Moradia': Array.from({length:12}, () => ({realized:1500, planned:0, total:1500})) }, 
        'Variáveis Essenciais': { 'Supermercado': Array.from({length:12}, () => ({realized:800, planned:0, total:800})) }, 
        'Discricionárias': { 'Lazer': Array.from({length:12}, () => ({realized:400, planned:0, total:400})) } 
      },
      monthlyTotals: Array.from({length:12}, () => ({income:5000, expense:2700, balance:2300})),
      summary: { incomeRealized:0, incomePlanned:60000, expenseRealized:0, expensePlanned:32400, currentBalance:2300, projectedBalance:27600 }
    };
  }

  // ==========================================
`;

code = code.replace(/<\/script>\s*$/, planningJS + '\n</script>');
fs.writeFileSync('gas-project/Scripts.html', code);
console.log('Scripts updated');
