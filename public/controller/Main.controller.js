sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox"
], function (Controller, JSONModel, MessageToast, MessageBox) {
    "use strict";

    var SOURCE_CONFIG = {
        semarang: {
            url:    "/api/sales-orders?iv_werks=3000&iv_auart=ZOR2",
            label:  "Semarang",
            werks:  "3000",
            auart:  "ZOR2"
        },
        surabaya: {
            url:    "/api/sales-orders?iv_werks=2000&iv_auart=ZOR1",
            label:  "Surabaya",
            werks:  "2000",
            auart:  "ZOR1"
        }
    };

    return Controller.extend("so.app.controller.Main", {

        onInit: function () {
            this._allData         = [];
            this._selectedKunnr   = null;
            this._selectedName    = null;
            this._searchQuery     = "";
            this._currentSource   = "semarang";
            this._dataLoaded      = false;
            this._statusFilter    = "ALL";
            this._monthFilter     = "ALL";
            this._expandedCustomers = {};
            this._updating        = false; // flag untuk mencegah reentrancy

            this.getView().setModel(new JSONModel({ results: [] }));

            window.__soApp = this;

            this._tick();
            this._clockTimer = setInterval(this._tick.bind(this), 1000);
            this._loadData();
        },

        onExit: function () {
            clearInterval(this._clockTimer);
            window.__soApp = null;
        },

        _tick: function () {
            var el = this.byId("clockText");
            if (!el) return;
            var n = new Date();
            el.setText(
                n.toLocaleDateString("en-US", { weekday: "short", day: "numeric", month: "short", year: "numeric" })
                + " · " + n.toLocaleTimeString("en-US")
            );
        },

        onSourceChange: function (oEvent) {
            var idx = oEvent.getParameter("selectedIndex");
            this._currentSource = (idx === 1) ? "surabaya" : "semarang";

            this._selectedKunnr = null;
            this._selectedName  = null;
            this._searchQuery   = "";
            this._statusFilter  = "ALL";
            this._monthFilter   = "ALL";
            this._expandedCustomers = {};

            var sf = this.byId("searchField");
            if (sf) sf.setValue("");
            var statusCombo = this.byId("statusFilter");
            if (statusCombo) statusCombo.setSelectedKey("ALL");
            var monthCombo = this.byId("monthFilter");
            if (monthCombo) monthCombo.setSelectedKey("ALL");

            this._updateSelectedBadge();

            var cfg = SOURCE_CONFIG[this._currentSource];
            MessageToast.show("Loading data " + cfg.label + "...");
            this._loadData();
        },

        _removeLeadingZero: function (str) {
            if (!str) return "";
            if (/^\d+$/.test(str)) return str.replace(/^0+/, '') || "0";
            return str;
        },

        _usdNoDecimal: function (v) {
            var num = Number(v || 0);
            if (isNaN(num)) num = 0;
            if (num % 1 === 0) return "$" + num.toLocaleString("en-US");
            return "$" + num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        },

        _intFmt: function (v) {
            var num = Number(v || 0);
            if (isNaN(num)) num = 0;
            return Math.floor(num).toLocaleString("en-US");
        },

        _plainNumberFmt: function (v) {
            var num = parseFloat(v);
            if (isNaN(num)) return "0";
            return Math.floor(num).toString();
        },

        _parseDateFromErdat: function (dateString) {
            if (!dateString || dateString.length !== 8) return null;
            var year = parseInt(dateString.substring(0, 4), 10);
            var month = parseInt(dateString.substring(4, 6), 10) - 1;
            var day = parseInt(dateString.substring(6, 8), 10);
            var date = new Date(year, month, day);
            return isNaN(date) ? null : date;
        },

        _parseDateFromYYYYMMDD: function (dateString) {
            if (!dateString) return null;
            var parts = dateString.split('-');
            if (parts.length !== 3) return null;
            var date = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
            return isNaN(date) ? null : date;
        },

        _formatToDDMMYYYY: function (d) {
            if (!d || isNaN(d)) return "-";
            return String(d.getDate()).padStart(2, "0") + "-"
                 + String(d.getMonth() + 1).padStart(2, "0") + "-"
                 + d.getFullYear();
        },

        _getAgeInfo: function (orderDateObj) {
            if (!orderDateObj) return { text: "-", status: "None", daysLeft: -1 };
            var today = new Date();
            today.setHours(0, 0, 0, 0);
            var diffDays = Math.round((today - orderDateObj) / (1000 * 60 * 60 * 24));
            if (diffDays < 0) diffDays = 0;
            var status = "None";
            if (diffDays <= 14) status = "Pending";
            else if (diffDays <= 25) status = "Low";
            else if (diffDays <= 45) status = "Medium";
            else if (diffDays <= 60) status = "Warning";
            else status = "Prioritas";
            return { text: diffDays + " days", status: status, daysLeft: diffDays };
        },

        _loadData: function () {
            var that = this;
            var cfg = SOURCE_CONFIG[this._currentSource];
            var isFirstLoad = !this._dataLoaded;

            fetch(cfg.url)
                .then(function (r) { return r.json(); })
                .then(function (raw) {
                    console.log("[" + cfg.label + "] Raw data:", raw);
                    var arr = Array.isArray(raw) ? raw : [];
                    that._allData = arr.map(function (x) {
                        var qtyBalance = parseFloat(x.qty_balance2) || 0;
                        var price = parseFloat(x.netpr) || 0;
                        var total = qtyBalance * price;
                        var vbeln = (x.vbeln || "").replace(/^0+/, '') || "0";
                        var posnr = (x.posnr || "").replace(/^0+/, '') || "0";
                        var matnr = that._removeLeadingZero(x.matnr || "");

                        var bstnk = x.bstnk || "";
                        var kwmeng = parseFloat(x.kwmeng) || 0;
                        var qtyGi = parseFloat(x.qty_gi) || 0;
                        var qtyBalance2 = qtyBalance;
                        var otsdoRaw = x.otsdo || "0";
                        var otsdoNum = parseFloat(otsdoRaw.toString().replace(/,/g, '')) || 0;
                        var kalabRaw = x.kalab || "0";
                        var kalab2Raw = x.kalab2 || "0";
                        var kalabNum = parseFloat(kalabRaw.toString().replace(/,/g, '')) || 0;
                        var kalab2Num = parseFloat(kalab2Raw.toString().replace(/,/g, '')) || 0;

                        var orderDateObj = null;
                        var orderDateFmt = "-";
                        if (x.erdat) {
                            orderDateObj = that._parseDateFromErdat(x.erdat);
                            orderDateFmt = that._formatToDDMMYYYY(orderDateObj);
                        }

                        var deliveryDateObj = null;
                        var deliveryDateFmt = "-";
                        if (x.edatu) {
                            deliveryDateObj = that._parseDateFromYYYYMMDD(x.edatu);
                            deliveryDateFmt = that._formatToDDMMYYYY(deliveryDateObj);
                        }

                        var ageInfo = that._getAgeInfo(orderDateObj);

                        return {
                            id:              vbeln + "-" + posnr,
                            vbeln:           vbeln,
                            bstnk:           bstnk,
                            customer:        x.name1 || x.kunnr || "-",
                            kunnr:           x.kunnr || "",
                            matnr:           matnr,
                            item:            x.maktx || x.matnr || "-",
                            quantity:        qtyBalance,
                            kwmeng:          kwmeng,
                            qtyGi:           qtyGi,
                            qtyBalance2:     qtyBalance2,
                            otsdo:           otsdoNum,
                            otsdoFmt:        that._plainNumberFmt(otsdoNum),
                            kalab:           kalabNum,
                            kalab2:          kalab2Num,
                            price:           price,
                            totalValue:      total,
                            kwmengFmt:       that._intFmt(kwmeng),
                            qtyGiFmt:        that._intFmt(qtyGi),
                            qtyBalance2Fmt:  that._intFmt(qtyBalance2),
                            kalabFmt:        that._plainNumberFmt(kalabNum),
                            kalab2Fmt:       that._plainNumberFmt(kalab2Num),
                            totalFmt:        that._usdNoDecimal(total),
                            orderDate:       orderDateObj,
                            orderDateFmt:    orderDateFmt,
                            deliveryDate:    deliveryDateObj,
                            deliveryDateFmt: deliveryDateFmt,
                            hariBerjalan:    ageInfo.text,
                            status:          ageInfo.status,
                            daysLeft:        ageInfo.daysLeft
                        };
                    });

                    // Deduplicate by vbeln-posnr (id)
                    var seen = {};
                    that._allData = that._allData.filter(function (r) {
                        if (seen[r.id]) return false;
                        seen[r.id] = true;
                        return true;
                    });

                    var scrollEl = document.querySelector(".sapMPageEnableScrolling")
                                || document.querySelector(".sapMScrollContScroll")
                                || document.documentElement;
                    var savedScroll = isFirstLoad ? 0 : (scrollEl.scrollTop || 0);

                    that._updateKPI();
                    that._buildCustTable();
                    that._applyFilters();
                    that._updateAgingAnalysis(); // sekali saja

                    that._dataLoaded = true;

                    if (savedScroll > 0) {
                        window.requestAnimationFrame(function () {
                            scrollEl.scrollTop = savedScroll;
                        });
                    }

                    var footerEl = that.byId("footerInfo");
                    if (footerEl) {
                        footerEl.setText(
                            "Sales Dashboard · " + cfg.label
                            + " · Werks " + cfg.werks
                            + " · Auart " + cfg.auart
                        );
                    }

                    var upEl = that.byId("updateTime");
                    if (upEl) upEl.setText("Update: " + new Date().toLocaleTimeString("en-US"));

                    MessageToast.show("✅ Data " + cfg.label + " loaded (" + that._allData.length + " rows)");
                })
                .catch(function (error) {
                    console.error("Error loading data:", error);
                    MessageToast.show("❌ Failed to load data " + cfg.label);
                });
        },

        _updateKPI: function () {
            var data   = this._allData;
            var tv     = data.reduce(function (s, r) { return s + r.totalValue; }, 0);
            var uc     = new Set(data.map(function (r) { return r.kunnr; })).size;
            var soUnik = new Set(data.map(function (r) { return r.vbeln; })).size;
            var totalQtyAll = data.reduce(function (s, r) { return s + r.quantity; }, 0);

            var totalOtsDO = data.reduce(function (s, r) { return s + r.otsdo; }, 0);
            var soOtsDOSet = new Set();
            data.forEach(function (r) {
                if (r.otsdo > 0) soOtsDOSet.add(r.vbeln);
            });
            var soOtsDOCount = soOtsDOSet.size;

            var cards = [
                {
                    icon: "📦",
                    label: "Total Qty / SO",
                    value: Math.floor(totalQtyAll).toLocaleString("en-US"),
                    sub: soUnik + " SO"
                },
                {
                    icon: "🚚",
                    label: "Ots. DO",
                    value: Math.floor(totalOtsDO).toLocaleString("en-US"),
                    sub: soOtsDOCount + " SO"
                },
                {
                    icon: "💵",
                    label: "Total Value",
                    value: this._usdNoDecimal(tv),
                    sub: null
                },
                {
                    icon: "👥",
                    label: "Customers",
                    value: String(uc),
                    sub: null
                }
            ];

            var html = cards.map(function (c) {
                return '<div class="kpiCard">'
                     +   '<div class="kpiCardIcon">'  + c.icon  + '</div>'
                     +   '<div class="kpiCardLabel">' + c.label + '</div>'
                     +   '<div class="kpiCardValue">' + c.value + '</div>'
                     +   (c.sub ? '<div class="kpiCardSub">' + c.sub + '</div>' : '')
                     + '</div>';
            }).join("");

            var kpiEl = this.byId("kpiRow");
            if (kpiEl) kpiEl.setContent('<div class="kpiRow">' + html + '</div>');
        },

        toggleCustomerDetail: function (kunnr) {
            if (this._expandedCustomers[kunnr]) {
                delete this._expandedCustomers[kunnr];
            } else {
                this._expandedCustomers[kunnr] = true;
            }
            this._updateAgingAnalysis();
        },

        _updateAgingAnalysis: function () {
            if (this._updating) return;
            this._updating = true;

            var that = this;
            var data = this._allData;

            var q = this._searchQuery;
            var kunnr = this._selectedKunnr;
            var statusFilter = this._statusFilter;
            var monthFilter = this._monthFilter;
            var filtered = data.filter(function (r) {
                var matchCust = !kunnr || r.kunnr === kunnr;
                var matchSearch = !q
                    || r.id.toLowerCase().includes(q)
                    || r.customer.toLowerCase().includes(q)
                    || r.item.toLowerCase().includes(q)
                    || r.matnr.toLowerCase().includes(q)
                    || (r.bstnk && r.bstnk.toLowerCase().includes(q));
                var matchStatus = !statusFilter || statusFilter === "ALL" || r.status === statusFilter;
                var matchMonth = that._matchMonthFilter(r.orderDate, monthFilter);
                return matchCust && matchSearch && matchStatus && matchMonth;
            });

            var categories = [
                { name: "0-14d", min: 0, max: 14, color: "#10b981", label: "Pending" },
                { name: "15-25d", min: 15, max: 25, color: "#eab308", label: "Low" },
                { name: "26-45d", min: 26, max: 45, color: "#f97316", label: "Medium" },
                { name: "46-60d", min: 46, max: 60, color: "#ef4444", label: "Warning" },
                { name: ">60d", min: 61, max: Infinity, color: "#b91c1c", label: "Prioritas" }
            ];

            // Build per-customer aging data (gunakan Map untuk memastikan unik)
            var custMap = new Map();
            filtered.forEach(function (r) {
                if (!custMap.has(r.kunnr)) {
                    custMap.set(r.kunnr, {
                        name: r.customer,
                        kunnr: r.kunnr,
                        totalValue: 0,
                        totalQty: 0,
                        totalOtsDO: 0,
                        count: 0,
                        buckets: [0, 0, 0, 0, 0],
                        bucketQty: [0, 0, 0, 0, 0],
                        bucketCount: [0, 0, 0, 0, 0],
                        bucketSO: [new Set(), new Set(), new Set(), new Set(), new Set()],
                        bucketOtsDO: [new Set(), new Set(), new Set(), new Set(), new Set()]
                    });
                }
                var c = custMap.get(r.kunnr);
                c.totalValue += r.totalValue;
                c.totalQty += r.quantity;
                c.totalOtsDO += r.otsdo;
                c.count++;
                for (var i = 0; i < categories.length; i++) {
                    if (r.daysLeft >= categories[i].min && r.daysLeft <= categories[i].max) {
                        c.buckets[i] += r.totalValue;
                        c.bucketQty[i] += r.quantity;
                        c.bucketCount[i]++;
                        c.bucketSO[i].add(r.vbeln);
                        if (r.otsdo > 0) {
                            c.bucketOtsDO[i].add(r.vbeln);
                        }
                        break;
                    }
                }
            });

            var custList = Array.from(custMap.values()).sort(function (a, b) { return b.totalValue - a.totalValue; });

            var grandTotal = filtered.reduce(function (s, r) { return s + r.totalValue; }, 0);
            var grandBuckets = [0, 0, 0, 0, 0];
            filtered.forEach(function (r) {
                for (var i = 0; i < categories.length; i++) {
                    if (r.daysLeft >= categories[i].min && r.daysLeft <= categories[i].max) {
                        grandBuckets[i] += r.totalValue;
                        break;
                    }
                }
            });

            // Summary header
            var html = '<div style="margin-bottom:10px;padding:10px 12px;background:linear-gradient(135deg,#f0fdf4,#ecfdf5);border-radius:10px;border:1px solid #bbf7d0;">';
            html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">';
            html += '<span style="font-weight:700;font-size:13px;color:#166534;">Total: ' + that._usdNoDecimal(grandTotal) + '</span>';
            html += '<span style="font-size:11px;color:#6b7280;">' + custList.length + ' customers · ' + filtered.length + ' lines</span>';
            html += '</div>';
            html += '<div style="display:flex;height:8px;border-radius:99px;overflow:hidden;background:#e5e7eb;">';
            for (var gi = 0; gi < categories.length; gi++) {
                var gp = grandTotal > 0 ? (grandBuckets[gi] / grandTotal * 100) : 0;
                if (gp > 0) {
                    html += '<div style="width:' + gp + '%;background:' + categories[gi].color + ';" title="' + categories[gi].label + ': ' + gp.toFixed(1) + '%"></div>';
                }
            }
            html += '</div>';
            html += '<div style="display:flex;gap:8px;margin-top:6px;flex-wrap:wrap;">';
            for (var li = 0; li < categories.length; li++) {
                html += '<span style="font-size:9px;color:#6b7280;display:flex;align-items:center;gap:3px;">';
                html += '<span style="width:8px;height:8px;border-radius:2px;background:' + categories[li].color + ';display:inline-block;"></span>';
                html += categories[li].label;
                html += '</span>';
            }
            html += '</div></div>';

            // Per-customer rows
            custList.forEach(function (c) {
                var safeName = c.name.replace(/</g, "&lt;").replace(/>/g, "&gt;");
                var safeKunnr = c.kunnr.replace(/'/g, "\\'");
                var isActive = that._selectedKunnr === c.kunnr;
                var isExpanded = that._expandedCustomers[c.kunnr] === true;
                var maxDayBucket = 0;
                for (var bi = categories.length - 1; bi >= 0; bi--) {
                    if (c.buckets[bi] > 0) { maxDayBucket = bi; break; }
                }
                var urgencyColor = categories[maxDayBucket].color;

                html += '<div class="aging-card" style="' + (isActive ? 'border-color:' + urgencyColor + ';background:#fffbeb;' : '') + '">';
                
                // Header: kiri select customer, kanan expand/collapse
                html += '<div style="display:flex;justify-content:space-between;align-items:center;">';
                html += '<div style="flex:1;min-width:0;cursor:pointer;" onclick="window.__soApp.selectCustomer(\'' + safeKunnr + '\',\'' + safeName.replace(/'/g, "\\'") + '\')">';
                html += '<div style="display:flex;justify-content:space-between;align-items:center;">';
                html += '<div style="flex:1;min-width:0;">';
                html += '<div style="font-weight:600;font-size:12px;color:' + (isActive ? '#c8401a' : '#374151') + ';white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + (isActive ? '▶ ' : '') + safeName + '</div>';
                html += '<div style="font-size:10px;color:#9ca3af;">' + c.count + ' lines · Qty ' + Math.floor(c.totalQty).toLocaleString('en-US') + '</div>';
                html += '</div>';
                html += '<div style="text-align:right;margin-left:8px;">';
                html += '<div style="font-weight:700;font-size:12px;color:#111827;">' + that._usdNoDecimal(c.totalValue) + '</div>';
                html += '</div>';
                html += '</div>';
                html += '</div>';
                html += '<div style="cursor:pointer;padding:4px 8px;" onclick="window.__soApp.toggleCustomerDetail(\'' + safeKunnr + '\'); event.stopPropagation();">';
                html += isExpanded ? '▼' : '▶';
                html += '</div>';
                html += '</div>';

                // Stacked bar
                html += '<div style="display:flex;height:6px;border-radius:99px;overflow:hidden;background:#e5e7eb;margin-top:6px;">';
                for (var ci = 0; ci < categories.length; ci++) {
                    var cp = c.totalValue > 0 ? (c.buckets[ci] / c.totalValue * 100) : 0;
                    if (cp > 0) {
                        html += '<div style="width:' + cp + '%;background:' + categories[ci].color + ';" title="' + categories[ci].label + ': ' + that._usdNoDecimal(c.buckets[ci]) + '"></div>';
                    }
                }
                html += '</div>';

                // Bucket chips (ringkasan) — hanya tampil jika TIDAK expanded
                if (!isExpanded) {
                    html += '<div style="display:flex;gap:4px;margin-top:5px;flex-wrap:wrap;">';
                    for (var di = 0; di < categories.length; di++) {
                        if (c.buckets[di] > 0) {
                            html += '<span style="font-size:9px;padding:1px 6px;border-radius:10px;background:' + categories[di].color + '20;color:' + categories[di].color + ';font-weight:600;">' + categories[di].label + ' ' + that._usdNoDecimal(c.buckets[di]) + '</span>';
                        }
                    }
                    html += '</div>';
                }

                // Expanded detail table (hanya jika isExpanded)
                if (isExpanded) {
                    html += '<div style="max-height:200px; overflow-y:auto; margin-top:8px;">';
                    html += '<table style="width:100%;border-collapse:collapse;font-size:11px;">';
                    html += '<thead><tr style="border-bottom:1px solid #e5e7eb;">';
                    html += '<th style="text-align:left;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;">Status</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;">SO</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;">Ots DO</th>';
                    html += '<th style="text-align:right;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;">Qty</th>';
                    html += '<th style="text-align:right;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;">Value</th>';
                    html += '</tr></thead><tbody>';
                    for (var ei = 0; ei < categories.length; ei++) {
                        if (c.bucketCount[ei] > 0) {
                            html += '<tr style="border-bottom:1px solid #f3f4f6;">';
                            html += '<td style="padding:4px 6px;"><span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:8px;height:8px;border-radius:2px;background:' + categories[ei].color + ';display:inline-block;"></span><span style="font-weight:700;font-size:12px;color:#1f2937;">' + categories[ei].label + '</span><span style="color:#6b7280;font-size:10px;">(' + categories[ei].name + ')</span></span></td>';
                            html += '<td style="padding:4px 6px;text-align:center;font-weight:600;">' + c.bucketSO[ei].size + '</td>';
                            html += '<td style="padding:4px 6px;text-align:center;font-weight:600;color:#e65100;">' + c.bucketOtsDO[ei].size + '</td>';
                            html += '<td style="padding:4px 6px;text-align:right;">' + Math.floor(c.bucketQty[ei]).toLocaleString('en-US') + '</td>';
                            html += '<td style="padding:4px 6px;text-align:right;font-weight:700;font-size:12px;color:#1f2937;">' + that._usdNoDecimal(c.buckets[ei]) + '</td>';
                            html += '</tr>';
                        }
                    }
                    var totalSO = new Set();
                    var totalOtsDOSet = new Set();
                    for (var si = 0; si < 5; si++) {
                        c.bucketSO[si].forEach(function(v){ totalSO.add(v); });
                        c.bucketOtsDO[si].forEach(function(v){ totalOtsDOSet.add(v); });
                    }
                    html += '<tr style="border-top:2px solid #d1d5db;background:#f9fafb;">';
                    html += '<td style="padding:4px 6px;font-weight:700;color:#374151;">Total</td>';
                    html += '<td style="padding:4px 6px;text-align:center;font-weight:700;">' + totalSO.size + '</td>';
                    html += '<td style="padding:4px 6px;text-align:center;font-weight:700;color:#e65100;">' + totalOtsDOSet.size + '</td>';
                    html += '<td style="padding:4px 6px;text-align:right;font-weight:700;">' + Math.floor(c.totalQty).toLocaleString('en-US') + '</td>';
                    html += '<td style="padding:4px 6px;text-align:right;font-weight:700;color:#166534;">' + that._usdNoDecimal(c.totalValue) + '</td>';
                    html += '</tr>';
                    html += '</tbody></table>';
                    html += '</div>';
                }
                html += '</div>';
            });

            if (custList.length === 0) {
                html += '<div style="text-align:center;padding:20px;color:#9ca3af;font-size:12px;">No data</div>';
            }

            var agingEl = this.byId("agingAnalysisHtml");
            if (agingEl) {
                // Suppress invalidation agar SAPUI5 tidak rerender ulang (mencegah duplikasi)
                agingEl.setProperty("content", html, true);
                var domRef = agingEl.getDomRef();
                if (domRef) {
                    domRef.innerHTML = html;
                }
            }
            this._updating = false;
        },

        _buildCustTable: function () {
            var C    = {};
            var that = this;

            this._allData.forEach(function (r) {
                if (!C[r.kunnr]) C[r.kunnr] = { name: r.customer, v: 0, c: 0, qty: 0, soSet: new Set() };
                C[r.kunnr].v   += r.totalValue;
                C[r.kunnr].c++;
                C[r.kunnr].qty += r.quantity;
                C[r.kunnr].soSet.add(r.vbeln);
            });

            var sorted = Object.entries(C).sort(function (a, b) { return b[1].v - a[1].v; });

            var th = function (txt, right) {
                return '<th style="padding:8px 10px;font-size:11px;font-weight:700;text-transform:uppercase;'
                     + 'letter-spacing:.04em;color:#6b7280;border-bottom:1px solid #e5e7eb;'
                     + 'text-align:' + (right ? "right" : "left") + '">' + txt + '</th>';
            };

            var rows = sorted.map(function (e, i) {
                var kunnr    = e[0];
                var d        = e[1];
                var isActive = that._selectedKunnr === kunnr;
                var safeName  = d.name.replace(/</g, "&lt;").replace(/>/g, "&gt;");
                var safeKunnr = kunnr.replace(/'/g, "\\'");
                var rowClass = isActive ? 'active-row' : '';

                return '<tr class="' + rowClass + '" onclick="window.__soApp.selectCustomer(\'' + safeKunnr + '\',\'' + safeName + '\')">'
                     + '<td style="padding:8px 10px;font-size:12px;color:#9ca3af">' + (i + 1) + '</td>'
                     + '<td style="padding:8px 10px;font-weight:600;font-size:13px;color:' + (isActive ? "#c8401a" : "#111827") + ';word-break:break-word;white-space:normal;">' + (isActive ? '▶ ' : '') + safeName + '</td>'
                     + '<td style="padding:8px 10px;color:#6b7280;font-size:12px;text-align:right">' + Math.floor(d.qty).toLocaleString('en-US') + '</td>'
                     + '<td style="padding:8px 10px;font-size:12px;font-weight:700;text-align:center;color:#4f46e5">' + d.soSet.size + '</td>'
                     + '<td style="padding:8px 10px;text-align:right;font-weight:700;font-size:13px;color:#111827">' + that._usdNoDecimal(d.v) + '</td>'
                     + '</tr>';
            }).join("");

            var clearBtn = this._selectedKunnr
                ? '<button onclick="window.__soApp.clearCustomer()" style="float:right;margin-bottom:6px;padding:3px 10px;font-size:11px;border:1px solid #d1d5db;border-radius:6px;background:#fff;cursor:pointer;color:#374151">✕ Show All</button>'
                : '';

            var html = '<div class="custTableContainer">' + clearBtn
                     + '<table class="custTable">'
                     + '<thead><tr>' + th("#") + th("Customer") + th("Total Qty", true) + th("Total SO", true) + th("Value (USD)", true) + '</tr></thead>'
                     + '<tbody>' + rows + '</tbody>'
                     + '</table></div>';
            var custHtml = this.byId("custTableHtml");
            if (custHtml) {
                custHtml.setProperty("content", html, true);
                var domRef = custHtml.getDomRef();
                if (domRef) {
                    domRef.innerHTML = html;
                }
            }
        },

        selectCustomer: function (kunnr, name) {
            if (this._selectedKunnr === kunnr) {
                this.clearCustomer();
                return;
            }
            this._selectedKunnr = kunnr;
            this._selectedName  = name;
            this._searchQuery   = "";
            var sf = this.byId("searchField");
            if (sf) sf.setValue("");

            this._updateSelectedBadge();
            this._buildCustTable();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        clearCustomer: function () {
            this._selectedKunnr = null;
            this._selectedName  = null;
            this._updateSelectedBadge();
            this._buildCustTable();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        _updateSelectedBadge: function () {
            var el = this.byId("selectedBadge");
            if (!el) return;
            if (this._selectedKunnr) {
                var safe = (this._selectedName || "").replace(/</g, "&lt;").replace(/>/g, "&gt;");
                el.setContent('<div style="display:inline-flex;align-items:center;gap:6px;background:#fff7ed;border:1px solid #fed7aa;border-radius:20px;padding:3px 12px;"><span style="font-size:11px;font-weight:600;color:#9a3412;">📋 SO Details :</span><span style="font-size:12px;font-weight:700;color:#c8401a;">' + safe + '</span></div>');
            } else {
                el.setContent('<div style="display:inline-flex;align-items:center;gap:6px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:20px;padding:3px 12px;"><span style="font-size:11px;font-weight:600;color:#0369a1;">📋 SO Details :</span><span style="font-size:12px;font-weight:700;color:#0284c7;">All Customers</span></div>');
            }
        },

        formatSalesOrderWithBstnk: function (bstnk, soId, showCustomerBelow, customerName) {
            var safeBstnk = (bstnk && bstnk !== "") ? bstnk.replace(/</g, "&lt;").replace(/>/g, "&gt;") : "";
            var safeSo = (soId || "-").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            var html = '<div style="display:flex;flex-direction:column;">';
            if (safeBstnk) {
                html += '<span style="font-size:10px; color:#6c757d;">' + safeBstnk + '</span>';
            }
            html += '<span style="font-weight:600;">' + safeSo + '</span>';
            if (showCustomerBelow) {
                var safeCust = (customerName || "-").replace(/</g, "&lt;").replace(/>/g, "&gt;");
                html += '<span style="font-size:10px; color:#6c757d;">' + safeCust + '</span>';
            }
            html += '</div>';
            return html;
        },

        onSearch: function (oEvent) {
            this._searchQuery = oEvent.getSource().getValue().toLowerCase();
            this._applyFilters();
        },

        onStatusFilterChange: function (oEvent) {
            this._statusFilter = oEvent.getParameter("selectedItem").getKey();
            this._applyFilters();
        },

        onMonthFilterChange: function (oEvent) {
            this._monthFilter = oEvent.getParameter("selectedItem").getKey();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        _matchMonthFilter: function (orderDate, monthKey) {
            if (!monthKey || monthKey === "ALL") return true;
            if (!orderDate) return false;
            return orderDate.getMonth() === parseInt(monthKey, 10);
        },

        _applyFilters: function () {
            var q     = this._searchQuery;
            var kunnr = this._selectedKunnr;
            var statusFilter = this._statusFilter;
            var monthFilter = this._monthFilter;
            var that = this;

            var filtered = this._allData.filter(function (r) {
                var matchCust   = !kunnr || r.kunnr === kunnr;
                var matchSearch = !q
                    || r.id.toLowerCase().includes(q)
                    || r.customer.toLowerCase().includes(q)
                    || r.item.toLowerCase().includes(q)
                    || r.matnr.toLowerCase().includes(q)
                    || (r.bstnk && r.bstnk.toLowerCase().includes(q));
                var matchStatus = !statusFilter || statusFilter === "ALL" || r.status === statusFilter;
                var matchMonth = that._matchMonthFilter(r.orderDate, monthFilter);
                return matchCust && matchSearch && matchStatus && matchMonth;
            });

            filtered.sort(function(a, b) {
                return (b.daysLeft || 0) - (a.daysLeft || 0);
            });

            var showCust = (this._selectedKunnr === null);
            filtered.forEach(function(item) {
                item.showCustomerBelow = showCust;
            });

            this.getView().getModel().setProperty("/results", filtered);
            var rowEl = this.byId("rowCount");
            if (rowEl) rowEl.setText(filtered.length + " rows");
        },

        _getActiveData: function () {
            var q     = this._searchQuery;
            var kunnr = this._selectedKunnr;
            var statusFilter = this._statusFilter;
            var monthFilter = this._monthFilter;
            var that = this;

            var filtered = this._allData.filter(function (r) {
                var matchCust   = !kunnr || r.kunnr === kunnr;
                var matchSearch = !q
                    || r.id.toLowerCase().includes(q)
                    || r.customer.toLowerCase().includes(q)
                    || r.item.toLowerCase().includes(q)
                    || r.matnr.toLowerCase().includes(q)
                    || (r.bstnk && r.bstnk.toLowerCase().includes(q));
                var matchStatus = !statusFilter || statusFilter === "ALL" || r.status === statusFilter;
                var matchMonth = that._matchMonthFilter(r.orderDate, monthFilter);
                return matchCust && matchSearch && matchStatus && matchMonth;
            });

            filtered.sort(function(a, b) {
                return (b.daysLeft || 0) - (a.daysLeft || 0);
            });

            return filtered;
        },

        _loadScript: function (url) {
            return new Promise(function (resolve, reject) {
                if (document.querySelector('script[src="' + url + '"]')) {
                    resolve(); return;
                }
                var s = document.createElement("script");
                s.src = url;
                s.onload  = function () { resolve(); };
                s.onerror = function () { reject(new Error("Failed to load: " + url)); };
                document.head.appendChild(s);
            });
        },

        _ensureLibs: function () {
            var that = this;
            var jobs = [];

            if (!window.XLSX) {
                jobs.push(that._loadScript("/libs/xlsx.full.min.js").then(function () {
                    if (typeof XLSX !== "undefined") window.XLSX = XLSX;
                }));
            }

            if (!window.jsPDF) {
                jobs.push(
                    that._loadScript("/libs/jspdf.umd.min.js").then(function () {
                        if (window.jspdf && window.jspdf.jsPDF) window.jsPDF = window.jspdf.jsPDF;
                    }).then(function () {
                        return that._loadScript("/libs/jspdf.plugin.autotable.min.js");
                    })
                );
            }

            return Promise.all(jobs);
        },

        onExportExcel: function () {
            var that = this;
            MessageToast.show("⏳ Preparing Excel file...");
            this._ensureLibs().then(function () {
                var XLSXLib = window.XLSX;
                if (!XLSXLib) {
                    MessageToast.show("❌ Excel library failed to load.");
                    return;
                }

                var cfg      = SOURCE_CONFIG[that._currentSource];
                var data     = that._getActiveData();
                var label    = that._selectedName ? that._selectedName : "All_Customers";
                var fileName = "SO_" + cfg.label + "_" + label.replace(/\s+/g, "_") + "_" + that._dateStamp() + ".xlsx";

                var wsData = [[
                    "Sales Order (PO)", "Created Date", "Material", "Qty Order", "Shipped", "Ots. SO",
                    "Ots. DO", "WHFG", "Packing", "Total Value (USD)", "Req. Delivery", "Days Running", "Status"
                ]];

                data.forEach(function (r) {
                    var salesOrderText = (r.bstnk ? r.bstnk + " / " : "") + r.id;
                    wsData.push([
                        salesOrderText,
                        r.orderDateFmt,
                        r.matnr + " - " + r.item,
                        Math.floor(r.kwmeng),
                        Math.floor(r.qtyGi),
                        Math.floor(r.qtyBalance2),
                        Math.floor(r.otsdo),
                        Math.floor(r.kalab).toString(),
                        Math.floor(r.kalab2).toString(),
                        r.totalValue,
                        r.deliveryDateFmt,
                        r.hariBerjalan,
                        r.status
                    ]);
                });

                var tv = data.reduce(function (s, r) { return s + r.totalValue; }, 0);
                wsData.push(["", "", "", "", "", "", "", "", "TOTAL", tv, "", "", ""]);

                var ws = XLSXLib.utils.aoa_to_sheet(wsData);
                ws["!cols"] = [
                    { wch: 25 }, { wch: 14 }, { wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
                    { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 10 }
                ];

                var wb = XLSXLib.utils.book_new();
                XLSXLib.utils.book_append_sheet(wb, ws, "Sales Orders");
                XLSXLib.writeFile(wb, fileName);

                MessageToast.show("✅ Excel file downloaded.");
            }).catch(function (e) {
                MessageToast.show("❌ Failed to load library: " + e.message);
            });
        },

        onExportPdf: function () {
            var that = this;
            MessageToast.show("⏳ Preparing PDF file...");
            this._ensureLibs().then(function () {
                var jsPDF = window.jsPDF || (window.jspdf && window.jspdf.jsPDF) || null;

                if (!jsPDF) {
                    MessageToast.show("❌ PDF library failed to load.");
                    return;
                }

                var cfg = SOURCE_CONFIG[that._currentSource];
                var data = that._getActiveData();
                var isAllCustomers = (that._selectedKunnr === null);
                var customerLabel = that._selectedName ? that._selectedName : "All Customers";

                var totalValue = data.reduce(function (sum, it) { return sum + it.totalValue; }, 0);
                var totalQty = data.reduce(function (sum, it) { return sum + it.quantity; }, 0);
                var totalSO = new Set(data.map(function (it) { return it.vbeln; })).size;

                var fileName = "SO_" + cfg.label + "_" + customerLabel.replace(/\s+/g, "_") + "_" + that._dateStamp() + ".pdf";

                var doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

                doc.setFillColor(134, 179, 130);
                doc.rect(0, 0, 297, 28, "F");
                doc.setTextColor(255, 255, 255);
                doc.setFontSize(14);
                doc.setFont("helvetica", "bold");
                doc.text("Sales Order Report — " + cfg.label, 14, 11);

                doc.setFontSize(9);
                doc.setFont("helvetica", "normal");
                doc.text("Customer: " + customerLabel, 14, 18);

                doc.setFontSize(8);
                doc.text("Generated: " + new Date().toLocaleString("en-US"), 200, 11);

                doc.setFontSize(9);
                doc.setFont("helvetica", "bold");
                doc.setTextColor(255, 255, 255);
                doc.text("Total Value: " + that._usdNoDecimal(totalValue), 14, 24);
                doc.text("Total SO: " + totalSO, 80, 24);
                doc.text("Total Qty: " + Math.floor(totalQty).toLocaleString("en-US"), 140, 24);

                doc.setTextColor(0, 0, 0);
                doc.setFontSize(8);
                doc.setFont("helvetica", "normal");

                var rows = data.map(function (r, idx) {
                    var salesOrderText = (r.bstnk ? r.bstnk + "\n" : "") + r.id;
                    if (isAllCustomers) {
                        salesOrderText = (r.bstnk ? r.bstnk + "\n" : "") + r.id + "\n" + r.customer;
                    }
                    return [
                        idx + 1,
                        salesOrderText,
                        r.orderDateFmt,
                        r.matnr + "\n" + r.item,
                        Math.floor(r.kwmeng),
                        Math.floor(r.qtyGi),
                        Math.floor(r.qtyBalance2),
                        Math.floor(r.otsdo),
                        Math.floor(r.kalab).toString(),
                        Math.floor(r.kalab2).toString(),
                        that._usdNoDecimal(r.totalValue),
                        r.deliveryDateFmt,
                        r.hariBerjalan,
                        r.status
                    ];
                });

                doc.autoTable({
                    startY: 32,
                    head: [[
                        "No", "Sales Order", "Created Date", "Material", "Qty Order", "Shipped", "Ots. SO",
                        "Ots. DO", "WHFG", "Packing", "Total Value", "Req. Delivery", "Days Running", "Status"
                    ]],
                    body: rows,
                    margin: { left: 10, right: 10 },
                    tableWidth: doc.internal.pageSize.getWidth() - 20,
                    styles: { fontSize: 7, cellPadding: 1.5, overflow: 'linebreak', textColor: [0,0,0], halign: 'center' },
                    headStyles: { fillColor: [134, 179, 130], textColor: 255, fontStyle: "bold", fontSize: 7, halign: 'center' },
                    alternateRowStyles: { fillColor: [240, 248, 240] },
                    columnStyles: {
                        0: { halign: "center", cellWidth: 8 },
                        1: { halign: "left", cellWidth: 36 },
                        2: { halign: "center", cellWidth: 18 },
                        3: { halign: "left", cellWidth: 42 },
                        4: { halign: "center", cellWidth: 16 },
                        5: { halign: "center", cellWidth: 16 },
                        6: { halign: "center", cellWidth: 16 },
                        7: { halign: "center", cellWidth: 15 },
                        8: { halign: "center", cellWidth: 16 },
                        9: { halign: "center", cellWidth: 16 },
                        10: { halign: "center", fontStyle: "bold", cellWidth: 22 },
                        11: { halign: "center", cellWidth: 18 },
                        12: { halign: "center", cellWidth: 18 },
                        13: { halign: "center", cellWidth: 20 }
                    },
                    didDrawPage: function (d) {
                        doc.setFontSize(7);
                        doc.setTextColor(100);
                        doc.text("Sales Dashboard - " + cfg.label + " - Page " + d.pageNumber, 14, doc.internal.pageSize.height - 5);
                    }
                });

                doc.save(fileName);
                MessageToast.show("✅ PDF file downloaded.");
            }).catch(function (e) {
                MessageToast.show("❌ Failed to load library: " + e.message);
            });
        },

        formatMaterial: function (matnr, maktx) {
            var code = (matnr || "-").replace(/</g, "&lt;");
            var desc = (maktx || "-").replace(/</g, "&lt;");
            return '<div style="display:flex;flex-direction:column;line-height:1.35">'
                 + '<span class="materialCode">' + code + '</span>'
                 + '<span class="materialDesc">' + desc + '</span>'
                 + '</div>';
        },

        formatStatus: function (status) {
            var statusClass = "";
            var statusText  = "";
            switch (status) {
                case "Pending": statusClass = "statusPending"; statusText = "Pending"; break;
                case "Low": statusClass = "statusLow"; statusText = "Low"; break;
                case "Medium": statusClass = "statusMedium"; statusText = "Medium"; break;
                case "Warning": statusClass = "statusWarning"; statusText = "⚠️ Warning"; break;
                case "Prioritas": statusClass = "statusPrioritas"; statusText = "🔥 Prioritas"; break;
                default: statusClass = "statusNone"; statusText = "-";
            }
            return '<span class="statusCell ' + statusClass + '">' + statusText + '</span>';
        },

        _dateStamp: function () {
            var d = new Date();
            return d.getFullYear()
                 + String(d.getMonth() + 1).padStart(2, "0")
                 + String(d.getDate()).padStart(2, "0");
        },

        onAgentOpen: function() {
            var dlg = this.byId("dlgAgent");
            if (dlg) {
                dlg.open();
                this.byId("agentInput").setValue("");
            }
        },

        onAgentClose: function() {
            this.byId("dlgAgent").close();
        },

        onAgentProcess: function() {
            var input = this.byId("agentInput").getValue().trim();
            if (!input) {
                MessageToast.show("Please enter a question.");
                return;
            }

            var data = this._allData;
            if (!data || data.length === 0) {
                MessageBox.show("Data not available yet. Please wait for data to load.", { title: "Agent" });
                this.byId("dlgAgent").close();
                return;
            }

            var lowerInput = input.toLowerCase();
            var isTotalQuery = lowerInput.includes("total") && (lowerInput.includes("value") || lowerInput.includes("nilai"));
            var filters = {
                customer: null,
                status: null,
                statusOutstanding: false,
                startDate: null,
                endDate: null
            };

            var knownCustomers = [...new Set(data.map(d => d.customer.toLowerCase()))];
            var matchedCustomer = null;
            for (var i = 0; i < knownCustomers.length; i++) {
                if (lowerInput.includes(knownCustomers[i])) {
                    matchedCustomer = knownCustomers[i];
                    break;
                }
            }
            if (matchedCustomer) {
                filters.customer = matchedCustomer;
            } else {
                var regex = /(?:customer|for|from)\s+([a-z0-9\s\.&]+?)(?:\s+(?:week|month|day|status|with|$))/i;
                var match = lowerInput.match(regex);
                if (match && match[1]) {
                    var possible = match[1].trim();
                    var found = knownCustomers.find(c => c.includes(possible));
                    if (found) filters.customer = found;
                }
            }

            if (lowerInput.includes("pending")) filters.status = "Pending";
            else if (lowerInput.includes("low")) filters.status = "Low";
            else if (lowerInput.includes("medium")) filters.status = "Medium";
            else if (lowerInput.includes("warning")) filters.status = "Warning";
            else if (lowerInput.includes("prioritas")) filters.status = "Prioritas";
            else if (lowerInput.includes("outstanding")) filters.statusOutstanding = true;

            var now = new Date();
            var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            if (lowerInput.includes("today")) {
                filters.startDate = today;
                filters.endDate = today;
            } else if (lowerInput.includes("this week") || lowerInput.includes("minggu ini")) {
                var startOfWeek = new Date(today);
                var day = today.getDay();
                var diff = (day === 0 ? 6 : day - 1);
                startOfWeek.setDate(today.getDate() - diff);
                filters.startDate = startOfWeek;
                filters.endDate = new Date(startOfWeek);
                filters.endDate.setDate(startOfWeek.getDate() + 6);
            } else if (lowerInput.includes("this month") || lowerInput.includes("bulan ini")) {
                filters.startDate = new Date(today.getFullYear(), today.getMonth(), 1);
                filters.endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
            }

            var filtered = data.filter(function(item) {
                if (filters.customer && !item.customer.toLowerCase().includes(filters.customer)) return false;
                if (filters.status && item.status !== filters.status) return false;
                if (filters.statusOutstanding && (item.status === "None" || item.status === "Pending")) return false;
                if (filters.startDate && item.orderDate) {
                    var d = new Date(item.orderDate.getFullYear(), item.orderDate.getMonth(), item.orderDate.getDate());
                    if (d < filters.startDate) return false;
                }
                if (filters.endDate && item.orderDate) {
                    var d2 = new Date(item.orderDate.getFullYear(), item.orderDate.getMonth(), item.orderDate.getDate());
                    if (d2 > filters.endDate) return false;
                }
                return true;
            });

            var message = "";
            if (isTotalQuery) {
                var totalValue = filtered.reduce(function(sum, it) { return sum + it.totalValue; }, 0);
                var totalQty = filtered.reduce(function(sum, it) { return sum + it.quantity; }, 0);
                message = "💰 **Total value:** " + this._usdNoDecimal(totalValue) + "\n";
                message += "📦 **Total qty:** " + Math.floor(totalQty).toLocaleString("en-US") + "\n";
                message += "📋 **Number of SO:** " + filtered.length;
            } else {
                if (filtered.length === 0) {
                    message = "No data matches your question.";
                } else {
                    message = "📊 **Filtered results (" + filtered.length + " items):**\n";
                    var displayItems = filtered.slice(0, 10);
                    displayItems.forEach(function(it) {
                        message += "• SO " + it.id + " | " + it.customer + " | " + it.totalFmt + " | " + it.status + "\n";
                    });
                    if (filtered.length > 10) {
                        message += "\n... and " + (filtered.length - 10) + " more items.";
                    }
                }
            }

            MessageBox.show(message, { title: "AI Sales Agent", icon: MessageBox.Icon.INFORMATION });
            this.byId("dlgAgent").close();
        }
    });
});