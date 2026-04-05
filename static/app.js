function setStatus(text, type) {
    const status = document.getElementById("status");
    status.className = "status " + type;
    status.textContent = "当前状态：" + text;
}

function bindFileName(inputId, textId) {
    const input = document.getElementById(inputId);
    const text = document.getElementById(textId);

    input.addEventListener("change", function () {
        const file = input.files[0];
        text.textContent = file ? file.name : "未选择文件";
    });
}

bindFileName("data_file", "data_file_name");
bindFileName("template_file", "template_file_name");

function renderSummary(containerId, products, customerRows) {
    const container = document.getElementById(containerId);

    if (!customerRows || customerRows.length === 0) {
        container.innerHTML = '<div class="summary-empty">暂无数据</div>';
        return;
    }

    const productTotals = {};
    products.forEach(p => productTotals[p] = 0);

    customerRows.forEach(row => {
        products.forEach(p => {
            productTotals[p] += Number(row[p] || 0);
        });
    });

    const total = customerRows.reduce((sum, row) => sum + Number(row["合计"] || 0), 0);

    let html = "";

    products.forEach(p => {
        html += `
            <div class="summary-item">
                <div class="summary-label">${p}</div>
                <div class="summary-value">${productTotals[p]}台</div>
                <div class="summary-note">单产品销量</div>
            </div>
        `;
    });

    html += `
        <div class="summary-item total">
            <div class="summary-label">合计</div>
            <div class="summary-value">${total}台</div>
            <div class="summary-note">全部产品总销量</div>
        </div>
    `;

    container.innerHTML = html;
}

function cellClass(value) {
    const num = Number(value || 0);
    return num === 0 ? "zero-value" : "";
}

function renderTable(containerId, rows, products, isStore = false) {
    const container = document.getElementById(containerId);

    if (!rows || rows.length === 0) {
        container.innerHTML = '<div class="table-placeholder">暂无数据</div>';
        return;
    }

    const headers = isStore
        ? ["门店名称", "导购", ...products, "合计"]
        : ["客户名称", ...products, "合计"];

    const tableClass = isStore ? "store-table" : "customer-table";

    let html = `<div class="table-wrap"><table class="${tableClass}"><thead><tr>`;
    headers.forEach(h => {
        html += `<th>${h}</th>`;
    });
    html += "</tr></thead><tbody>";

    rows.forEach(row => {
        html += "<tr>";

        if (isStore) {
            html += `<td title="${row["门店名称"] || ""}">${row["门店名称"] || ""}</td>`;
            html += `<td title="${row["导购"] || ""}">${row["导购"] || ""}</td>`;
        } else {
            html += `<td title="${row["客户名称"] || ""}">${row["客户名称"] || ""}</td>`;
        }

        products.forEach(p => {
            const value = row[p] ?? 0;
            html += `<td class="${cellClass(value)}">${value}</td>`;
        });

        const total = row["合计"] ?? 0;
        html += `<td>${total}</td>`;
        html += "</tr>";
    });

    html += "</tbody></table></div>";
    container.innerHTML = html;
}

function renderDownloads(resultFile, logFile) {
    const box = document.getElementById("downloads");
    box.innerHTML = `
        <a class="download-btn" href="/download/${resultFile}" target="_blank">输出</a>
        <a class="download-btn" href="/download/${logFile}" target="_blank">运行日志</a>
    `;
}

document.getElementById("runBtn").addEventListener("click", async function () {
    const runBtn = document.getElementById("runBtn");

    const dataFile = document.getElementById("data_file").files[0];
    const templateFile = document.getElementById("template_file").files[0];

    const product1 = document.getElementById("product1").value.trim();
    const product2 = document.getElementById("product2").value.trim();
    const product3 = document.getElementById("product3").value.trim();
    const product4 = document.getElementById("product4").value.trim();

    if (!dataFile) {
        setStatus("请先上传微服务原表", "error");
        return;
    }

    if (!templateFile) {
        setStatus("请先上传模板文件", "error");
        return;
    }

    if (!product1) {
        setStatus("产品1必填", "error");
        return;
    }

    const formData = new FormData();
    formData.append("data_file", dataFile);
    formData.append("template_file", templateFile);
    formData.append("product1", product1);
    formData.append("product2", product2);
    formData.append("product3", product3);
    formData.append("product4", product4);

    runBtn.disabled = true;
    runBtn.textContent = "生成中";
    setStatus("处理中", "processing");

    try {
        const res = await fetch("/process", {
            method: "POST",
            body: formData
        });

        const data = await res.json();

        if (!data.ok) {
            setStatus("处理失败：" + data.msg, "error");
            runBtn.disabled = false;
            runBtn.textContent = "开始生成";
            return;
        }

        setStatus("处理完成", "success");
        renderSummary("summary", data.products, data.customer_rows);
        renderTable("customer_preview", data.customer_rows, data.products, false);
        renderTable("store_preview", data.store_rows, data.products, true);
        renderDownloads(data.result_file, data.log_file);

    } catch (e) {
        setStatus("请求失败：" + e, "error");
    }

    runBtn.disabled = false;
    runBtn.textContent = "开始生成";
});