let expenses = [];
let categories = ["仕込み", "当日", "ホール料金", "その他"];
let editIndex = null;

const categorySelect = document.getElementById("category");
const tableArea = document.getElementById("tableArea");
const budgetInput = document.getElementById("budgetInput");

function formatYen(num) {
  return Number(num).toLocaleString("ja-JP");
}

function getBudget() {
  return Number(budgetInput.value) || 0;
}

function renderCategories() {
  categorySelect.innerHTML = "";
  categories.forEach(cat => {
    const option = document.createElement("option");
    option.value = cat;
    option.textContent = cat;
    categorySelect.appendChild(option);
  });
}

function renderTable() {
  tableArea.innerHTML = "";
  let totalAll = 0;

  categories.forEach(cat => {
    const filtered = expenses
      .map((e, i) => ({ ...e, index: i }))
      .filter(e => e.category === cat);

    if (filtered.length === 0) return;

    const block = document.createElement("div");
    block.className = "table-block";

    const title = document.createElement("div");
    title.className = "table-title";
    title.textContent = `■ ${cat}`;
    block.appendChild(title);

    const table = document.createElement("table");

    table.innerHTML = `
      <tr>
        <th>項目</th>
        <th>金額</th>
        <th>備考</th>
        <th></th>
      </tr>
    `;

    let total = 0;

    filtered.forEach(e => {
      total += Number(e.amount);

      table.innerHTML += `
        <tr>
          <td>${e.item}</td>
          <td>${formatYen(e.amount)}</td>
          <td>${e.note}</td>
          <td>
            <button onclick="editExpense(${e.index})">編集</button>
            <button onclick="deleteExpense(${e.index})">削除</button>
          </td>
        </tr>
      `;
    });

    totalAll += total;

    table.innerHTML += `
      <tr>
        <td><strong>計</strong></td>
        <td><strong>${formatYen(total)}</strong></td>
        <td></td>
        <td></td>
      </tr>
    `;

    block.appendChild(table);
    tableArea.appendChild(block);
  });

  document.getElementById("remaining").textContent =
    formatYen(getBudget() - totalAll);
}

function editExpense(index) {
  const e = expenses[index];
  document.getElementById("category").value = e.category;
  document.getElementById("item").value = e.item;
  document.getElementById("amount").value = e.amount;
  document.getElementById("note").value = e.note;
  editIndex = index;
}

function deleteExpense(index) {
  if (!confirm("削除する？戻せないぞ？")) return;
  expenses.splice(index, 1);
  renderTable();
}

document.getElementById("addExpenseBtn").addEventListener("click", () => {
  const category = categorySelect.value;
  const item = document.getElementById("item").value.trim();
  const amount = document.getElementById("amount").value;
  const note = document.getElementById("note").value.trim();

  if (!item || !amount) {
    alert("項目と金額を入力してください");
    return;
  }

  if (editIndex !== null) {
    expenses[editIndex] = { category, item, amount, note };
    editIndex = null;
  } else {
    expenses.push({ category, item, amount, note });
  }

  document.getElementById("item").value = "";
  document.getElementById("amount").value = "";
  document.getElementById("note").value = "";

  renderTable();
});


// 🔥 区分追加（重複防止つき）
document.getElementById("addCategoryBtn").addEventListener("click", () => {
  const name = prompt("新しい区分名を入力してください");
  if (!name || !name.trim()) return;

  const trimmed = name.trim();

  const exists = categories.some(
    cat => cat.toLowerCase() === trimmed.toLowerCase()
  );

  if (exists) {
    alert("その区分はもう存在するで😎");
    return;
  }

  categories.push(trimmed);
  renderCategories();
});

budgetInput.addEventListener("input", renderTable);


/* ===========================
   🔥 Excel出力
=========================== */

document.getElementById("exportExcelBtn").addEventListener("click", () => {

  let ws_data = [];
  let totalAll = 0;
  let rowIndex = 0;

  categories.forEach(cat => {
    const filtered = expenses.filter(e => e.category === cat);
    if (filtered.length === 0) return;

    ws_data.push([`■ ${cat}`, "", ""]);
    rowIndex++;

    ws_data.push(["項目", "金額", "備考"]);
    rowIndex++;

    let total = 0;

    filtered.forEach(e => {
      total += Number(e.amount);
      ws_data.push([e.item, Number(e.amount), e.note]);
      rowIndex++;
    });

    ws_data.push(["計", total, ""]);
    rowIndex++;

    ws_data.push(["", "", ""]);
    rowIndex++;

    totalAll += total;
  });

  ws_data.push(["総合計", totalAll, ""]);
  ws_data.push(["予算", getBudget(), ""]);
  ws_data.push(["残り", getBudget() - totalAll, ""]);

  const ws = XLSX.utils.aoa_to_sheet(ws_data);

  ws["!cols"] = [
    { wch: 28 },
    { wch: 15 },
    { wch: 38 }
  ];

  ws["!rows"] = ws_data.map(() => ({ hpt: 20 }));

  const range = XLSX.utils.decode_range(ws["!ref"]);

  function styleCell(cell, style) {
    if (!cell.s) cell.s = {};
    Object.assign(cell.s, style);
  }

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = 0; C <= 2; ++C) {

      const ref = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[ref];
      if (!cell) continue;

      styleCell(cell, {
        font: { name: "Meiryo", sz: 11 },
        alignment: { vertical: "center" },
        border: {
          top: { style: "hair" },
          bottom: { style: "hair" },
          left: { style: "hair" },
          right: { style: "hair" }
        }
      });

      if (C === 1 && typeof cell.v === "number") {
        cell.z = '"¥"#,##0';
        styleCell(cell, {
          alignment: { horizontal: "right", vertical: "center" }
        });
      }

      const value = cell.v;

      if (value === "項目" || value === "金額" || value === "備考") {
        styleCell(cell, {
          fill: { patternType: "solid", fgColor: { rgb: "DDEBF7" } },
          font: { bold: true }
        });
      }

      if (value === "計") {
        styleCell(cell, { font: { bold: true } });
        const amountCell = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
        if (amountCell) styleCell(amountCell, { font: { bold: true } });
      }

      if (value === "総合計" || value === "予算") {
        styleCell(cell, {
          fill: { patternType: "solid", fgColor: { rgb: "DDEBF7" } },
          font: { bold: true }
        });

        const amountCell = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
        if (amountCell) {
          styleCell(amountCell, {
            fill: { patternType: "solid", fgColor: { rgb: "DDEBF7" } },
            font: { bold: true }
          });
        }
      }

      if (value === "残り") {
        styleCell(cell, {
          fill: { patternType: "solid", fgColor: { rgb: "FFF2CC" } },
          font: { bold: true }
        });

        const amountCell = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
        if (amountCell) {
          styleCell(amountCell, {
            fill: { patternType: "solid", fgColor: { rgb: "FFF2CC" } },
            font: { bold: true }
          });
        }
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "会計データ");
  XLSX.writeFile(wb, "会計データ.xlsx");
});

renderCategories();
renderTable();
