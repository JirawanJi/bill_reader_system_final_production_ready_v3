document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("uploadForm");
    const resultBox = document.getElementById("result");

    if (!form || !resultBox) return;

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        const billType = form.dataset.billType;
        const fileInput = document.getElementById("files");

        if (!fileInput || !fileInput.files.length) {
            showResult("กรุณาเลือกไฟล์ PDF", "error");
            return;
        }

        const formData = new FormData();
        for (const file of fileInput.files) {
            formData.append("files", file);
        }

        showResult("กำลังประมวลผล...", "info");

        try {
            const response = await fetch(`/upload/${billType}`, {
                method: "POST",
                body: formData
            });

            const data = await response.json();
            console.log("UPLOAD RESPONSE =", data);

            if (!response.ok || !data.success) {
                let html = `<h3>ไม่สำเร็จ</h3><p>${data.message || "เกิดข้อผิดพลาด"}</p>`;

                if (data.errors && data.errors.length) {
                    html += "<ul>";
                    data.errors.forEach(err => {
                        html += `<li>${escapeHtml(String(err))}</li>`;
                    });
                    html += "</ul>";
                }

                showResult(html, "error", true);
                return;
            }

            let html = `
                <h3>สำเร็จ</h3>
                <p>${escapeHtml(String(data.message || ""))}</p>
                <p><a class="btn btn-primary" href="${data.download_url}">ดาวน์โหลดไฟล์ Excel</a></p>
            `;

            if (data.rows && data.rows.length) {
                html += `
                    <div class="table-wrap">
                        <table>
                            <thead>
                                <tr>
                                    <th>File</th>
                                    <th>Invoice No</th>
                                    <th>Date</th>
                                    <th>Amount</th>
                                    <th>Store ID</th>
                                    <th>Cost Center</th>
                                    <th>Profit Center</th>
                                </tr>
                            </thead>
                            <tbody>
                `;

                data.rows.forEach(row => {
                    console.log("ROW DEBUG =", row);

                    const fileName =
                        row.source_file ||
                        row.filename ||
                        row.original_filename ||
                        "";

                    const invoiceNo =
                        row.reference ||
                        row.invoice_no ||
                        row.invoice ||
                        "";

                    const invoiceDate =
                        row.invoice_date ||
                        row.date ||
                        "";

                    const amount =
                        row.amount ??
                        row.total_amount ??
                        "";

                    const storeId =
                        row.store_id ||
                        row.store ||
                        "";

                    const costCenter =
                        row.cost_center ||
                        row.costcenter ||
                        row.CostCenter ||
                        "";

                    const profitCenter =
                        row.profit_center ||
                        row.profitcenter ||
                        row.ProfitCenter ||
                        "";

                    html += `
                        <tr>
                            <td>${escapeHtml(String(fileName))}</td>
                            <td>${escapeHtml(String(invoiceNo))}</td>
                            <td>${escapeHtml(String(invoiceDate))}</td>
                            <td>${escapeHtml(String(amount))}</td>
                            <td>${escapeHtml(String(storeId))}</td>
                            <td>${escapeHtml(String(costCenter))}</td>
                            <td>${escapeHtml(String(profitCenter))}</td>
                        </tr>
                    `;
                });

                html += `
                            </tbody>
                        </table>
                    </div>
                `;
            }

            showResult(html, "success", true);

        } catch (error) {
            console.error("UPLOAD ERROR =", error);
            showResult(`เกิดข้อผิดพลาด: ${error.message}`, "error");
        }
    });

    function showResult(message, type = "info", isHtml = false) {
        resultBox.classList.remove("hidden", "success", "error", "info");
        resultBox.classList.add(type);

        if (isHtml) {
            resultBox.innerHTML = message;
        } else {
            resultBox.textContent = message;
        }
    }

    function escapeHtml(text) {
        return text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }
});