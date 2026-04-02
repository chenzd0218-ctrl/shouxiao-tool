function highlight(text) {
    if (!text) return "";

    let html = text;

    html = html.replace(/(合计[：:]\s*)(\d+)(\s*台)?/g, '$1<span class="sum">$2</span>$3');
    html = html.replace(/([：:]\s*)(\d+)(\s*台)/g, '$1<span class="num">$2</span>$3');

    return html;
}

function showToast(message) {
    const toast = document.getElementById("toast");
    toast.innerText = message;
    toast.classList.add("show");
    clearTimeout(window.__toastTimer);
    window.__toastTimer = setTimeout(() => {
        toast.classList.remove("show");
    }, 1800);
}

function updateFileName(inputId, outputId) {
    const input = document.getElementById(inputId);
    const output = document.getElementById(outputId);
    if (!input || !output) return;

    const file = input.files && input.files[0];
    output.innerText = file ? file.name : "未选择文件";
}

async function run() {
    const dataFile = document.getElementById("dataFile");
    const templateFile = document.getElementById("templateFile");
    const launchDate = document.getElementById("launchDate");
    const p1 = document.getElementById("p1");
    const p2 = document.getElementById("p2");
    const summary = document.getElementById("summary");
    const reportDay = document.getElementById("report_day");
    const report5Day = document.getElementById("report_5day");
    const btn = document.getElementById("runBtn");

    if (!dataFile.files[0]) {
        showToast("请先上传微服务原表");
        dataFile.focus();
        return;
    }

    if (!templateFile.files[0]) {
        showToast("请先上传模板文件");
        templateFile.focus();
        return;
    }

    if (!launchDate.value) {
        showToast("请选择首销日期");
        launchDate.focus();
        return;
    }

    const formData = new FormData();
    formData.append("data_file", dataFile.files[0]);
    formData.append("template_file", templateFile.files[0]);
    formData.append("launch_date", launchDate.value);
    formData.append("p1", p1.value);
    formData.append("p2", p2.value);

    btn.disabled = true;
    btn.innerText = "生成中...";
    summary.innerText = "处理中，请稍候...";
    reportDay.innerHTML = "";
    report5Day.innerHTML = "";

    try {
        const res = await fetch("/process", {
            method: "POST",
            body: formData
        });

        const data = await res.json();

        if (data.ok) {
            summary.innerText = data.summary || "完成";
            reportDay.innerHTML = highlight(data.report_day || "");
            report5Day.innerHTML = highlight(data.report_5day || "");
            showToast("生成成功");
            reportDay.scrollIntoView({ behavior: "smooth", block: "start" });
        } else {
            summary.innerText = data.msg || "处理失败";
            showToast("处理失败");
        }
    } catch (e) {
        summary.innerText = "失败：" + e;
        showToast("请求失败");
    } finally {
        btn.disabled = false;
        btn.innerText = "生成首销通报";
    }
}

function copy(id) {
    const text = document.getElementById(id).innerText;
    navigator.clipboard.writeText(text);
    showToast("已复制");
}

function copyAll() {
    const day = document.getElementById("report_day").innerText || "";
    const d5 = document.getElementById("report_5day").innerText || "";
    const text = `${day}\n\n${d5}`.trim();
    navigator.clipboard.writeText(text);
    showToast("已复制全部通报");
}

document.addEventListener("DOMContentLoaded", function () {
    const dataFile = document.getElementById("dataFile");
    const templateFile = document.getElementById("templateFile");

    if (dataFile) {
        dataFile.addEventListener("change", function () {
            updateFileName("dataFile", "dataFileName");
        });
    }

    if (templateFile) {
        templateFile.addEventListener("change", function () {
            updateFileName("templateFile", "templateFileName");
        });
    }
});