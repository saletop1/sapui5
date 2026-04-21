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
            this._reqDeliveryFilter = "ALL";
            this._expandedCustomers = {};
            this._agingTypeFilter = "ALL";
            this._updating        = false;
            this._currentPage     = 1;
            this._pageSize        = 50;
            this._filteredData    = [];
            this._pickerYear      = new Date().getFullYear();
            this._pickerFilterType = null;

            this.getView().setModel(new JSONModel({ results: [], resultsPaged: [] }));

            window.__soApp = this;

            this._darkMode = localStorage.getItem("soDarkMode") === "true";
            if (this._darkMode) document.body.classList.add("dark-theme");

            var that = this;
            this.getView().addEventDelegate({
                onAfterRendering: function () {
                    that._updateThemeIcon();
                }
            });

            this._tick();
            this._clockTimer = setInterval(this._tick.bind(this), 1000);
            this._loadData();
        },

        onToggleTheme: function () {
            this._darkMode = !this._darkMode;
            document.body.classList.toggle("dark-theme", this._darkMode);
            localStorage.setItem("soDarkMode", String(this._darkMode));
            this._updateThemeIcon();
        },

        _updateThemeIcon: function () {
            var btn = this.byId("themeToggle");
            if (btn) {
                btn.setText(this._darkMode ? "☀️" : "🌙");
                btn.setTooltip(this._darkMode ? "Switch to Light Mode" : "Switch to Dark Mode");
            }
        },

        onExit: function () {
            clearInterval(this._clockTimer);
            window.__soApp = null;
        },

        _tick: function () {
            var n = new Date();
            var dateTime = n.toLocaleDateString("en-US", { weekday: "short", day: "numeric", month: "short", year: "numeric" })
                + " · " + n.toLocaleTimeString("en-US", { hour12: false });
            var dateOnly = n.toLocaleDateString("en-US", { weekday: "long", day: "numeric", month: "long", year: "numeric" });

            var el = this.byId("clockText");
            if (el) el.setText(dateTime);

            var footerClock = this.byId("footerClock");
            if (footerClock) footerClock.setText(dateOnly);
        },

        onSourceChange: function (oEvent) {
            var idx = oEvent.getParameter("selectedIndex");
            this._currentSource = (idx === 1) ? "surabaya" : "semarang";
            this._hideMonthPicker();

            this._selectedKunnr = null;
            this._selectedName  = null;
            this._searchQuery   = "";
            this._statusFilter  = "ALL";
            this._monthFilter   = "ALL";
            this._reqDeliveryFilter = "ALL";
            this._agingTypeFilter   = "ALL";
            this._expandedCustomers = {};

            this._resetFilterControls();

            this._updateSelectedBadge();

            var cfg = SOURCE_CONFIG[this._currentSource];
            MessageToast.show("Loading data " + cfg.label + "...");
            this._loadData();
        },

        /** Helper: ambil elemen DOM di dalam core:HTML wrapper */
        _innerEl: function (id, selector) {
            var ctrl = this.byId(id);
            var dom  = ctrl && ctrl.getDomRef ? ctrl.getDomRef() : null;
            return dom ? (selector ? dom.querySelector(selector) : dom.firstElementChild) : null;
        },

        /** Reset semua filter controls ke nilai awal */
        _resetFilterControls: function () {
            var statusCtrl = this.byId("statusFilter");
            if (statusCtrl) statusCtrl.setSelectedKey("ALL");
            var searchCtrl = this.byId("searchField");
            if (searchCtrl) searchCtrl.setValue("");
            this._updatePickerBtnLabel("monthFilterBtn", "ALL");
            this._updatePickerBtnLabel("reqDeliveryFilterBtn", "ALL");
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

        /**
         * Clipboard copy — dipanggil dari onclick di formatter HTML.
         * Fallback ke execCommand jika Clipboard API tidak tersedia.
         */
        _copyText: function (text, el) {
            var doFlash = function (ok) {
                if (!el) return;
                var orig = el.style.background;
                el.style.background = ok ? "#d1fae5" : "#fee2e2";
                el.style.transition = "background 0.3s";
                setTimeout(function () {
                    el.style.background = orig;
                    el.style.transition  = "";
                }, 700);
            };
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text).then(
                    function () { doFlash(true); },
                    function () { doFlash(false); }
                );
            } else {
                // Fallback: buat textarea sementara
                var ta = document.createElement("textarea");
                ta.value = text;
                ta.style.cssText = "position:fixed;top:-9999px;left:-9999px;opacity:0;";
                document.body.appendChild(ta);
                ta.select();
                try {
                    document.execCommand("copy");
                    doFlash(true);
                } catch (e) {
                    doFlash(false);
                }
                document.body.removeChild(ta);
            }
        },

        _fmtPct: function (v) {
            var num = Number(v || 0);
            if (isNaN(num)) num = 0;
            // Bulatkan max 2 desimal, hilangkan trailing zero: 100 → "100%", 34.56 → "34.56%", 34.5 → "34.5%"
            var rounded = parseFloat(num.toFixed(2));
            return rounded + "%";
        },

        /**
         * Resolve API "type" string ke kategori tampilan.
         *   contains WOOD + METAL → "W & M"
         *   contains WOOD only    → "Wood"
         *   contains METAL only   → "Metal"
         *   otherwise             → "" (uncategorised)
         */
        _resolveType: function (typeStr) {
            if (!typeStr) return "";
            var t        = typeStr.toUpperCase();
            var hasWood  = t.indexOf("WOOD")  !== -1;
            var hasMetal = t.indexOf("METAL") !== -1;
            if (hasWood && hasMetal) return "W & M";
            if (hasWood)             return "Wood";
            if (hasMetal)            return "Metal";
            return "";
        },

        /** Dipanggil oleh chip button di Aging Analysis type filter */
        onAgingTypeChange: function (val) {
            this._agingTypeFilter = val;
            this._applyFilters();
            this._updateAgingAnalysis();
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
            if (!orderDateObj) return { text: "-", status: "New", daysLeft: -1 };
            var today = new Date();
            today.setHours(0, 0, 0, 0);
            var diffDays = Math.round((today - orderDateObj) / (1000 * 60 * 60 * 24));
            if (diffDays < 0) diffDays = 0;
            var status = "New";
            if (diffDays <= 7) status = "New";
            else if (diffDays <= 14) status = "Pending";
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
                            type:            x.type || "",
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

                    that._populateMonthFilter();
                    that._populateReqDeliveryFilter();
                    that._updateKPI();
                    that._buildCustTable();
                    that._applyFilters();
                    that._updateAgingAnalysis();

                    that._dataLoaded = true;

                    if (savedScroll > 0) {
                        window.requestAnimationFrame(function () {
                            scrollEl.scrollTop = savedScroll;
                        });
                    }

                    var footerEl = that.byId("footerInfo");
                    if (footerEl) footerEl.setText("Sales Dashboard · " + cfg.label);

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
                    sub: soUnik + " SO",
                    gradient: "kpiGrad1"
                },
                {
                    icon: "🚚",
                    label: "Ots. DO",
                    value: Math.floor(totalOtsDO).toLocaleString("en-US"),
                    sub: soOtsDOCount + " SO",
                    gradient: "kpiGrad2"
                },
                {
                    icon: "💵",
                    label: "Total Value",
                    value: this._usdNoDecimal(tv),
                    sub: null,
                    gradient: "kpiGrad3"
                },
                {
                    icon: "👥",
                    label: "Customers",
                    value: String(uc),
                    sub: null,
                    gradient: "kpiGrad4"
                }
            ];

            var html = cards.map(function (c) {
                return '<div class="kpiCard ' + c.gradient + '">'
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

        filterByBucket: function (kunnr, statusLabel) {
            // Select customer
            var custData = this._allData.find(function(r) { return r.kunnr === kunnr; });
            var custName = custData ? custData.customer : "";
            this._selectedKunnr = kunnr;
            this._selectedName = custName;

            // Set status filter
            this._statusFilter = statusLabel;
            var statusCtrl = this.byId("statusFilter");
            if (statusCtrl) statusCtrl.setSelectedKey(statusLabel);

            this._searchQuery = "";
            var searchCtrl = this.byId("searchField");
            if (searchCtrl) searchCtrl.setValue("");

            this._updateSelectedBadge();
            this._buildCustTable();
            this._applyFilters();
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
            var agingTypeFilter = this._agingTypeFilter || "ALL";
            var filtered = data.filter(function (r) {
                var matchCust = !kunnr || r.kunnr === kunnr;
                var matchSearch = !q
                    || r.id.toLowerCase().includes(q)
                    || r.customer.toLowerCase().includes(q)
                    || r.item.toLowerCase().includes(q)
                    || r.matnr.toLowerCase().includes(q)
                    || (r.bstnk && r.bstnk.toLowerCase().includes(q));
                var matchStatus = !statusFilter || statusFilter === "ALL" || r.status === statusFilter;
                var matchMonth = that._matchMonthFilter(r.deliveryDate, monthFilter);
                var matchReqDel = that._matchReqDeliveryFilter(r.deliveryDate);
                var matchType  = agingTypeFilter === "ALL" || that._resolveType(r.type) === agingTypeFilter;
                return matchCust && matchSearch && matchStatus && matchMonth && matchReqDel && matchType;
            });

            var categories = [
                { name: "0-7d", min: 0, max: 7, color: "#06b6d4", label: "New" },
                { name: "8-14d", min: 8, max: 14, color: "#10b981", label: "Pending" },
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
                        buckets: [0, 0, 0, 0, 0, 0],
                        bucketQty: [0, 0, 0, 0, 0, 0],
                        bucketOtsDOQty: [0, 0, 0, 0, 0, 0],
                        bucketCount: [0, 0, 0, 0, 0, 0],
                        bucketSO: [new Set(), new Set(), new Set(), new Set(), new Set(), new Set()]
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
                        c.bucketOtsDOQty[i] += r.otsdo;
                        c.bucketCount[i]++;
                        c.bucketSO[i].add(r.vbeln);
                        break;
                    }
                }
            });

            var custList = Array.from(custMap.values()).sort(function (a, b) { return b.totalValue - a.totalValue; });

            var grandTotal = filtered.reduce(function (s, r) { return s + r.totalValue; }, 0);
            var grandBuckets = [0, 0, 0, 0, 0, 0];
            filtered.forEach(function (r) {
                for (var i = 0; i < categories.length; i++) {
                    if (r.daysLeft >= categories[i].min && r.daysLeft <= categories[i].max) {
                        grandBuckets[i] += r.totalValue;
                        break;
                    }
                }
            });

            // ── Type filter chips — dirender ke agingTypeFilterBar (di LUAR ScrollContainer) ──
            var typeKeys   = ["ALL", "Metal", "Wood", "W & M"];
            var typeLabels = { "ALL": "All", "Metal": "Metal", "Wood": "Wood", "W & M": "W&M" };
            var chips = typeKeys.map(function (k) {
                var isActive = agingTypeFilter === k;
                var baseStyle = 'display:inline-flex;align-items:center;height:22px;padding:0 9px;'
                              + 'border-radius:11px;font-size:10px;font-weight:' + (isActive ? '700' : '500') + ';'
                              + 'cursor:pointer;white-space:nowrap;border:1px solid;'
                              + 'transition:all 0.15s;font-family:inherit;outline:none;';
                var colorStyle = isActive
                    ? 'background:#16a34a;color:#fff;border-color:#16a34a;'
                    : 'background:transparent;color:var(--text-muted,#6b7280);border-color:var(--border-medium,#c0c4cc);';
                return '<button onclick="window.__soApp.onAgingTypeChange(\'' + k + '\')" '
                     + 'style="' + baseStyle + colorStyle + '">'
                     + typeLabels[k] + '</button>';
            }).join("");

            var fbHtml = '<div style="display:flex;align-items:center;gap:5px;'
                       + 'padding:5px 10px;border-bottom:1px solid var(--border-light,rgba(0,0,0,0.08));'
                       + 'background:var(--bg-glass-solid,#fff);">'
                       + '<span style="font-size:10px;font-weight:600;color:var(--text-muted,#6b7280);'
                       + 'white-space:nowrap;margin-right:2px;">Type:</span>'
                       + chips
                       + '</div>';

            var filterBarEl = this.byId("agingTypeFilterBar");
            if (filterBarEl) {
                filterBarEl.setProperty("content", fbHtml, true);
                var fbDom = filterBarEl.getDomRef();
                if (fbDom) fbDom.innerHTML = fbHtml;
            }

            // Summary header
            var html = '<div class="agingSummary">';
            html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">';
            html += '<span style="font-weight:700;font-size:13px;color:#166534;">Total: ' + that._usdNoDecimal(grandTotal) + '</span>';
            html += '<span style="font-size:11px;color:var(--text-muted);">' + custList.length + ' customers · ' + filtered.length + ' lines</span>';
            html += '</div>';
            html += '<div style="display:flex;height:8px;border-radius:99px;overflow:hidden;background:var(--border-light);">';
            for (var gi = 0; gi < categories.length; gi++) {
                var gp = grandTotal > 0 ? (grandBuckets[gi] / grandTotal * 100) : 0;
                if (gp > 0) {
                    html += '<div style="width:' + gp + '%;background:' + categories[gi].color + ';" title="' + categories[gi].label + ': ' + gp.toFixed(1) + '%"></div>';
                }
            }
            html += '</div>';
            html += '<div style="display:flex;gap:8px;margin-top:6px;flex-wrap:wrap;">';
            for (var li = 0; li < categories.length; li++) {
                html += '<span style="font-size:9px;color:var(--text-muted);display:flex;align-items:center;gap:3px;">';
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
                            var safeKunnrChip = c.kunnr.replace(/'/g, "\\'");
                            html += '<span style="font-size:9px;padding:1px 6px;border-radius:10px;background:' + categories[di].color + '20;color:' + categories[di].color + ';font-weight:600;cursor:pointer;transition:opacity 0.15s;" onmouseover="this.style.opacity=\'0.7\'" onmouseout="this.style.opacity=\'1\'" onclick="event.stopPropagation();window.__soApp.filterByBucket(\'' + safeKunnrChip + '\',\'' + categories[di].label + '\')">' + categories[di].label + ' ' + that._usdNoDecimal(c.buckets[di]) + '</span>';
                        }
                    }
                    html += '</div>';
                }

                // Expanded detail table (hanya jika isExpanded)
                if (isExpanded) {
                    html += '<div style="max-height:200px; overflow-y:auto; margin-top:8px;">';
                    html += '<table style="width:100%;border-collapse:collapse;font-size:11px;">';
                    html += '<thead><tr style="border-bottom:1px solid #e5e7eb;">';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;white-space:nowrap;vertical-align:middle;">Status</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;white-space:normal;word-break:break-word;line-height:1.25;vertical-align:middle;">SO</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;white-space:normal;word-break:break-word;line-height:1.25;vertical-align:middle;">Qty</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;white-space:normal;word-break:break-word;line-height:1.25;vertical-align:middle;">Ots DO</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;white-space:normal;word-break:break-word;line-height:1.25;vertical-align:middle;">% DO</th>';
                    html += '<th style="text-align:center;padding:4px 6px;color:#6b7280;font-size:10px;font-weight:700;white-space:normal;word-break:break-word;line-height:1.25;vertical-align:middle;">Ots. SO Value</th>';
                    html += '</tr></thead><tbody>';
                    for (var ei = 0; ei < categories.length; ei++) {
                        if (c.bucketCount[ei] > 0) {
                            var safeKunnrBucket = c.kunnr.replace(/'/g, "\\'");
                            var bQty = c.bucketQty[ei];
                            var bOtsDO = c.bucketOtsDOQty[ei];
                            var doPct = bQty > 0 ? ((bQty - bOtsDO) / bQty * 100) : 0;
                            var doPctColor = doPct >= 80 ? '#16a34a' : doPct >= 50 ? '#ca8a04' : '#dc2626';
                            html += '<tr style="border-bottom:1px solid #f3f4f6;cursor:pointer;transition:background 0.15s;" onmouseover="this.style.background=\'#f0f9ff\'" onmouseout="this.style.background=\'transparent\'" onclick="window.__soApp.filterByBucket(\'' + safeKunnrBucket + '\',\'' + categories[ei].label + '\')">';
                            html += '<td style="padding:4px 6px;"><span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:8px;height:8px;border-radius:2px;background:' + categories[ei].color + ';display:inline-block;"></span><span style="font-weight:700;font-size:12px;color:#1f2937;">' + categories[ei].label + '</span><span style="color:#6b7280;font-size:10px;">(' + categories[ei].name + ')</span></span></td>';
                            html += '<td style="padding:4px 6px;text-align:center;font-weight:600;">' + c.bucketSO[ei].size + '</td>';
                            html += '<td style="padding:4px 6px;text-align:right;">' + Math.floor(bQty).toLocaleString('en-US') + '</td>';
                            html += '<td style="padding:4px 6px;text-align:right;font-weight:600;color:#e65100;">' + Math.floor(bOtsDO).toLocaleString('en-US') + '</td>';
                            html += '<td style="padding:4px 6px;text-align:center;font-weight:700;color:' + doPctColor + ';">' + that._fmtPct(doPct) + '</td>';
                            html += '<td style="padding:4px 6px;text-align:right;font-weight:700;font-size:12px;color:#1f2937;">' + that._usdNoDecimal(c.buckets[ei]) + '</td>';
                            html += '</tr>';
                        }
                    }
                    var totalSO = new Set();
                    for (var si = 0; si < 6; si++) {
                        c.bucketSO[si].forEach(function(v){ totalSO.add(v); });
                    }
                    var totalDOPct = c.totalQty > 0 ? ((c.totalQty - c.totalOtsDO) / c.totalQty * 100) : 0;
                    var totalDOPctColor = totalDOPct >= 80 ? '#16a34a' : totalDOPct >= 50 ? '#ca8a04' : '#dc2626';
                    html += '<tr style="border-top:2px solid #d1d5db;background:#f9fafb;">';
                    html += '<td style="padding:4px 6px;font-weight:700;color:#374151;">Total</td>';
                    html += '<td style="padding:4px 6px;text-align:center;font-weight:700;">' + totalSO.size + '</td>';
                    html += '<td style="padding:4px 6px;text-align:right;font-weight:700;">' + Math.floor(c.totalQty).toLocaleString('en-US') + '</td>';
                    html += '<td style="padding:4px 6px;text-align:right;font-weight:700;color:#e65100;">' + Math.floor(c.totalOtsDO).toLocaleString('en-US') + '</td>';
                    html += '<td style="padding:4px 6px;text-align:center;font-weight:700;color:' + totalDOPctColor + ';">' + that._fmtPct(totalDOPct) + '</td>';
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
                if (!C[r.kunnr]) C[r.kunnr] = { name: r.customer, v: 0, c: 0, qty: 0, otsDO: 0, soSet: new Set() };
                C[r.kunnr].v   += r.totalValue;
                C[r.kunnr].c++;
                C[r.kunnr].qty += r.quantity;
                C[r.kunnr].otsDO += r.otsdo;
                C[r.kunnr].soSet.add(r.vbeln);
            });

            var sorted = Object.entries(C).sort(function (a, b) { return b[1].v - a[1].v; });

            var th = function (txt, nowrap) {
                var wrapStyle = nowrap
                    ? 'white-space:nowrap;'
                    : 'white-space:normal;word-break:break-word;line-height:1.25;';
                return '<th style="padding:8px 10px;font-size:11px;font-weight:700;text-transform:uppercase;'
                     + 'letter-spacing:.04em;color:var(--text-muted);border-bottom:1px solid var(--border-medium);'
                     + wrapStyle
                     + 'vertical-align:middle;text-align:center;">' + txt + '</th>';
            };

            var rows = sorted.map(function (e, i) {
                var kunnr    = e[0];
                var d        = e[1];
                var isActive = that._selectedKunnr === kunnr;
                var safeName  = d.name.replace(/</g, "&lt;").replace(/>/g, "&gt;");
                var safeKunnr = kunnr.replace(/'/g, "\\'");
                var rowClass = isActive ? 'active-row' : '';

                var doPct = d.qty > 0 ? ((d.qty - d.otsDO) / d.qty * 100) : 0;
                var doPctColor = doPct >= 80 ? '#16a34a' : doPct >= 50 ? '#ca8a04' : '#dc2626';

                // Nama pendek = kata pertama saja (untuk mobile)
                var shortName = safeName.split(/[\s,]+/)[0];

                return '<tr class="' + rowClass + '" onclick="window.__soApp.selectCustomer(\'' + safeKunnr + '\',\'' + safeName + '\')">'
                     + '<td class="ctNo">' + (i + 1) + '</td>'
                     + '<td class="ctName" style="color:' + (isActive ? "#c8401a" : "var(--text-primary)") + ';">'
                     +   (isActive ? '▶ ' : '')
                     +   '<span class="ctNameFull">' + safeName + '</span>'
                     +   '<span class="ctNameShort">' + shortName + '</span>'
                     + '</td>'
                     + '<td class="ctNum">' + Math.floor(d.qty).toLocaleString('en-US') + '</td>'
                     + '<td class="ctNum" style="color:#e65100;font-weight:600;">' + Math.floor(d.otsDO).toLocaleString('en-US') + '</td>'
                     + '<td class="ctNum" style="color:' + doPctColor + ';font-weight:700;">' + that._fmtPct(doPct) + '</td>'
                     + '<td class="ctNum" style="color:#4f46e5;font-weight:700;">' + d.soSet.size + '</td>'
                     + '<td class="ctVal">' + that._usdNoDecimal(d.v) + '</td>'
                     + '</tr>';
            }).join("");

            var clearBtn = this._selectedKunnr
                ? '<button onclick="window.__soApp.clearCustomer()" style="float:right;margin-bottom:6px;padding:3px 10px;font-size:11px;border:1px solid var(--border-medium);border-radius:6px;background:var(--bg-card);cursor:pointer;color:var(--text-secondary)">✕ Show All</button>'
                : '';

            var html = '<div class="custTableContainer">' + clearBtn
                     + '<table class="custTable">'
                     + '<thead><tr>' + th("No", true) + th("Customer") + th("Total Qty") + th("Ots. DO") + th("% DO") + th("Ots. SO") + th("Ots. SO Value") + '</tr></thead>'
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
            var searchCtrl = this.byId("searchField");
            if (searchCtrl) searchCtrl.setValue("");

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
            var rawBstnk  = (bstnk  || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            var safeSo    = (soId   || "-").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            var rawSo     = (soId   || "-").replace(/\\/g, "\\\\").replace(/'/g, "\\'");

            var copyStyle = 'cursor:copy;border-radius:3px;padding:1px 2px;margin:-1px -2px;'
                          + 'transition:background 0.3s;';
            var html = '<div style="display:flex;flex-direction:column;line-height:1.4;">';
            if (safeBstnk) {
                html += '<span onclick="event.stopPropagation();window.__soApp._copyText(\'' + rawBstnk + '\',this)" '
                      + 'title="Click to copy" '
                      + 'style="font-size:10px;color:#6c757d;' + copyStyle + '">' + safeBstnk + '</span>';
            }
            html += '<span onclick="event.stopPropagation();window.__soApp._copyText(\'' + rawSo + '\',this)" '
                  + 'title="Click to copy" '
                  + 'style="font-weight:600;' + copyStyle + '">' + safeSo + '</span>';
            if (showCustomerBelow) {
                var safeCust = (customerName || "-").replace(/</g, "&lt;").replace(/>/g, "&gt;");
                html += '<span style="font-size:10px;color:#6c757d;">' + safeCust + '</span>';
            }
            html += '</div>';
            return html;
        },

        onSearch: function (oEvent) {
            var that = this;
            this._searchQuery = oEvent.getSource().getValue().toLowerCase();
            if (this._searchTimer) clearTimeout(this._searchTimer);
            this._searchTimer = setTimeout(function () {
                that._applyFilters();
            }, 300);
        },

        onStatusFilterChange: function (oEvent) {
            this._statusFilter = oEvent.getSource().getSelectedKey();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        /* =========================================================
         *  MONTH PICKER — pengganti dropdown Delivery & Req. Del
         * ========================================================= */

        _MONTH_SHORT: ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
        _MONTH_FULL:  ["January","February","March","April","May","June",
                       "July","August","September","October","November","December"],

        /** Hitung ISO week number dari sebuah Date */
        _isoWeekNum: function (date) {
            var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
            d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
            var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
            return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
        },

        /** Hitung week-to-month dan month-to-weeks mapping untuk satu tahun */
        _buildWeekMonthMap: function (year) {
            var jan4 = new Date(year, 0, 4);
            var dow  = jan4.getDay() || 7;                        // ISO Mon=1..Sun=7
            var mon  = new Date(jan4);
            mon.setDate(jan4.getDate() - dow + 1);                // Monday of W01

            var weekToMonth = {};
            var monthToWeeks = { 0:[],1:[],2:[],3:[],4:[],5:[],6:[],7:[],8:[],9:[],10:[],11:[] };

            for (var w = 1; w <= 53; w++) {
                var thu = new Date(mon); thu.setDate(mon.getDate() + 3);   // Thursday
                if (thu.getFullYear() !== year) break;
                var m = thu.getMonth();
                weekToMonth[w] = m;
                monthToWeeks[m].push(w);
                mon = new Date(mon); mon.setDate(mon.getDate() + 7);
            }
            return { weekToMonth: weekToMonth, monthToWeeks: monthToWeeks };
        },

        /** Update teks tombol picker (core:HTML)
         *  getDomRef() pada core:HTML mengembalikan elemen konten itu sendiri (si <button>),
         *  bukan wrapper div — jadi langsung pakai elemen tersebut. */
        _updatePickerBtnLabel: function (btnId, key) {
            var ctrl = this.byId(btnId);
            if (!ctrl) return;
            var el = ctrl.getDomRef ? ctrl.getDomRef() : null;
            if (!el) return;
            // el bisa berupa <button> langsung, atau wrapper yang mengandung <button>
            var btn = (el.classList && el.classList.contains("soPickerBtn"))
                ? el
                : el.querySelector(".soPickerBtn");
            if (!btn) return;

            var label;
            if (!key || key === "ALL") {
                label = "All \u25bc";
            } else if (key.charAt(0) === "W") {
                label = key + " \u25bc";
            } else {
                label = this._MONTH_SHORT[parseInt(key, 10)] + " \u25bc";
            }
            btn.textContent = label;
        },

        onDeliveryPickerOpen: function (el) {
            this._showMonthPicker(el, "delivery");
        },
        onReqDeliveryPickerOpen: function (el) {
            this._showMonthPicker(el, "reqdelivery");
        },

        _showMonthPicker: function (anchorEl, filterType) {
            var that = this;
            var existing = document.getElementById("__monthPickerPopup");
            if (existing) {
                var sameType = existing.dataset.filterType === filterType;
                this._hideMonthPicker();
                if (sameType) return;
            }

            this._pickerFilterType = filterType;
            this._pickerYear = this._pickerYear || new Date().getFullYear();

            var overlay = document.createElement("div");
            overlay.id = "__monthPickerOverlay";
            overlay.onclick = function () { that._hideMonthPicker(); };
            document.body.appendChild(overlay);

            var popup = document.createElement("div");
            popup.id = "__monthPickerPopup";
            popup.dataset.filterType = filterType;
            document.body.appendChild(popup);

            if (anchorEl) {
                var rect   = anchorEl.getBoundingClientRect();
                var popupW = 306;
                var popupH = 240;
                var left = Math.min(rect.left, window.innerWidth - popupW - 10);
                var top  = rect.bottom + 5;
                if (top + popupH > window.innerHeight) top = rect.top - popupH - 5;
                popup.style.left = Math.max(6, left) + "px";
                popup.style.top  = Math.max(6, top)  + "px";
            }

            this._renderPicker();
        },

        /** Render isi popup (dipanggil tiap ganti tahun) */
        _renderPicker: function () {
            var that     = this;
            var popup    = document.getElementById("__monthPickerPopup");
            if (!popup) return;

            var filterType = this._pickerFilterType;
            var yr         = this._pickerYear;
            var curKey     = filterType === "delivery" ? this._monthFilter : this._reqDeliveryFilter;
            var title      = filterType === "delivery" ? "Delivery Month" : "Req. Del Month";

            var map        = this._buildWeekMonthMap(yr);
            var w2m        = map.weekToMonth;   // {1:0, 2:0, ..., 52:11}
            var m2w        = map.monthToWeeks;  // {0:[1,2,3,4], ..., 11:[...]}
            var totalWeeks = Object.keys(w2m).length;

            var curIsWeek  = curKey && curKey.charAt(0) === "W";
            var curWeekNum = curIsWeek ? parseInt(curKey.substring(1), 10) : -1;
            var curMonthIdx = (!curIsWeek && curKey !== "ALL" && curKey !== null) ? parseInt(curKey, 10) : -1;

            var html = "";

            /* ── Top bar (compact): title | ‹ year › | All | ✕ ── */
            html += '<div class="mp-bar">';
            html += '<span class="mp-title">' + title + '</span>';
            html += '<button class="mp-nav-btn" onclick="window.__soApp._pickerPrevYear()">‹</button>';
            html += '<span class="mp-year-label">' + yr + '</span>';
            html += '<button class="mp-nav-btn" onclick="window.__soApp._pickerNextYear()">›</button>';
            html += '<button class="mp-all-sm" onclick="window.__soApp._pickerSelectAll()">All</button>';
            html += '<button class="mp-close-btn" onclick="window.__soApp._hideMonthPicker()">✕</button>';
            html += '</div>';

            /* ── Body: left = weeks (scrollable) | right = months grid ── */
            html += '<div class="mp-body">';

            /* Panel kiri: semua weeks W01–W52/53 (scrollable, clickable) */
            html += '<div class="mp-weeks-col" id="__mpWeeksCol">';
            for (var w = 1; w <= totalWeeks; w++) {
                var wMonth  = w2m[w];
                var isCurWk = (w === curWeekNum);
                var cls = "mp-week-cell" + (isCurWk ? " mp-week-sel" : "");
                var wLabel = "W" + String(w).padStart(2, "0");
                html += '<div class="' + cls + '" data-w="' + w + '" data-m="' + wMonth + '"'
                      + ' onmouseenter="window.__soApp._pickerHiWeek(' + w + ',true)"'
                      + ' onmouseleave="window.__soApp._pickerHiWeek(' + w + ',false)"'
                      + ' onclick="window.__soApp._pickerSelectWeek(' + w + ')">'
                      + wLabel + '</div>';
            }
            html += '</div>';

            /* Panel kanan: bulan 4 baris × 3 kolom */
            html += '<div class="mp-months-grid">';
            var monthRows = [[0,1,2],[3,4,5],[6,7,8],[9,10,11]];
            monthRows.forEach(function (row) {
                html += '<div class="mp-month-row">';
                row.forEach(function (m) {
                    var isSel  = (m === curMonthIdx);
                    var mWeeks = (m2w[m] || []).join(",");
                    var cls2   = "mp-month-cell" + (isSel ? " mp-month-sel" : "");
                    html += '<div class="' + cls2 + '" data-m="' + m + '" data-weeks="' + mWeeks + '"'
                          + ' onmouseenter="window.__soApp._pickerHiMonth(' + m + ',true)"'
                          + ' onmouseleave="window.__soApp._pickerHiMonth(' + m + ',false)"'
                          + ' onclick="window.__soApp._pickerSelectMonth(' + m + ')">'
                          + that._MONTH_SHORT[m] + '</div>';
                });
                html += '</div>';
            });
            html += '</div>';
            html += '</div>'; // .mp-body

            popup.innerHTML = html;

            /* Scroll weeks ke posisi yang relevan */
            var col = document.getElementById("__mpWeeksCol");
            if (col) {
                var target = curIsWeek
                    ? col.querySelector(".mp-week-sel")
                    : col.querySelector("[data-m='" + (curMonthIdx >= 0 ? curMonthIdx : new Date().getMonth()) + "']");
                if (!target) target = col.querySelector("[data-w='" + that._isoWeekNum(new Date()) + "']");
                if (target) { col.scrollTop = target.offsetTop - col.clientHeight / 2 + target.offsetHeight / 2; }
            }
        },

        /* Hover highlight: week → highlight matching months */
        _pickerHiWeek: function (w, on) {
            var popup = document.getElementById("__monthPickerPopup");
            if (!popup) return;
            var map = this._buildWeekMonthMap(this._pickerYear);
            var m   = map.weekToMonth[w];
            var el  = popup.querySelector(".mp-month-cell[data-m='" + m + "']");
            if (el) el.classList.toggle("mp-hi", on);
        },

        /* Hover highlight: month → highlight matching weeks */
        _pickerHiMonth: function (m, on) {
            var popup = document.getElementById("__monthPickerPopup");
            if (!popup) return;
            popup.querySelectorAll(".mp-week-cell[data-m='" + m + "']").forEach(function (el) {
                el.classList.toggle("mp-hi", on);
            });
        },

        _pickerPrevYear: function () {
            this._pickerYear = (this._pickerYear || new Date().getFullYear()) - 1;
            this._renderPicker();
        },
        _pickerNextYear: function () {
            this._pickerYear = (this._pickerYear || new Date().getFullYear()) + 1;
            this._renderPicker();
        },

        _pickerSelectWeek: function (w) {
            var key = "W" + w;
            if (this._pickerFilterType === "delivery") {
                this._monthFilter = key;
                this._updatePickerBtnLabel("monthFilterBtn", key);
            } else {
                this._reqDeliveryFilter = key;
                this._updatePickerBtnLabel("reqDeliveryFilterBtn", key);
            }
            this._hideMonthPicker();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        _pickerSelectMonth: function (monthIndex) {
            var key = String(monthIndex);
            if (this._pickerFilterType === "delivery") {
                this._monthFilter = key;
                this._updatePickerBtnLabel("monthFilterBtn", key);
            } else {
                this._reqDeliveryFilter = key;
                this._updatePickerBtnLabel("reqDeliveryFilterBtn", key);
            }
            this._hideMonthPicker();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        _pickerSelectAll: function () {
            if (this._pickerFilterType === "delivery") {
                this._monthFilter = "ALL";
                this._updatePickerBtnLabel("monthFilterBtn", "ALL");
            } else {
                this._reqDeliveryFilter = "ALL";
                this._updatePickerBtnLabel("reqDeliveryFilterBtn", "ALL");
            }
            this._hideMonthPicker();
            this._applyFilters();
            this._updateAgingAnalysis();
        },

        _hideMonthPicker: function () {
            var popup   = document.getElementById("__monthPickerPopup");
            var overlay = document.getElementById("__monthPickerOverlay");
            if (popup)   popup.remove();
            if (overlay) overlay.remove();
            this._pickerFilterType = null;
        },

        /* Lama — tidak dipakai lagi tapi dipertahankan agar tidak error jika dipanggil */
        onMonthFilterChange: function () {},
        onReqDeliveryFilterChange: function () {},

        _matchReqDeliveryFilter: function (deliveryDate) {
            var filter = this._reqDeliveryFilter;
            if (!filter || filter === "ALL") return true;
            if (!deliveryDate) return false;
            var dd = new Date(deliveryDate);
            if (filter.charAt(0) === "W") {
                return this._isoWeekNum(dd) === parseInt(filter.substring(1), 10);
            }
            return dd.getMonth() === parseInt(filter, 10);
        },

        _populateReqDeliveryFilter: function () {
            // Tidak dipakai — filter kini menggunakan custom month picker popup
        },

        _matchMonthFilter: function (deliveryDate, monthKey) {
            if (!monthKey || monthKey === "ALL") return true;
            if (!deliveryDate) return true;   // jika tidak ada tanggal, ikutkan saja
            var dd = new Date(deliveryDate);

            if (monthKey.charAt(0) === "W") {
                // Kumulatif s/d akhir week terpilih (Sunday ISO week)
                var wNum = parseInt(monthKey.substring(1), 10);
                var yr   = this._pickerYear || new Date().getFullYear();
                // Cari Monday of W01: Jan 4 selalu di W01
                var jan4 = new Date(yr, 0, 4);
                var dow  = jan4.getDay() || 7;
                var w1Mon = new Date(jan4);
                w1Mon.setDate(jan4.getDate() - dow + 1);
                // Monday of wNum
                var wMon = new Date(w1Mon);
                wMon.setDate(w1Mon.getDate() + (wNum - 1) * 7);
                // Sunday (akhir week) = Monday + 6
                var wSun = new Date(wMon);
                wSun.setDate(wMon.getDate() + 6);
                wSun.setHours(23, 59, 59, 999);
                return dd <= wSun;
            }

            // Kumulatif s/d akhir bulan terpilih
            var selectedMonth = parseInt(monthKey, 10);
            var yr2 = this._pickerYear || new Date().getFullYear();
            var endDate = new Date(yr2, selectedMonth + 1, 0, 23, 59, 59, 999);
            return dd <= endDate;
        },

        _populateMonthFilter: function () {
            // Tidak dipakai — filter kini menggunakan custom month picker popup
        },

        _applyFilters: function () {
            var q     = this._searchQuery;
            var kunnr = this._selectedKunnr;
            var statusFilter    = this._statusFilter;
            var monthFilter     = this._monthFilter;
            var agingTypeFilter = this._agingTypeFilter || "ALL";
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
                var matchMonth  = that._matchMonthFilter(r.deliveryDate, monthFilter);
                var matchReqDel = that._matchReqDeliveryFilter(r.deliveryDate);
                var matchType   = agingTypeFilter === "ALL" || that._resolveType(r.type) === agingTypeFilter;
                return matchCust && matchSearch && matchStatus && matchMonth && matchReqDel && matchType;
            });

            filtered.sort(function(a, b) {
                // Urutkan by customer dulu, lalu paling tua (daysLeft terbesar)
                var custCmp = (a.customer || "").localeCompare(b.customer || "");
                if (custCmp !== 0) return custCmp;
                return (b.daysLeft || 0) - (a.daysLeft || 0);
            });

            var showCust = (this._selectedKunnr === null);
            filtered.forEach(function(item) {
                item.showCustomerBelow = showCust;
            });

            this._filteredData = filtered;
            this._currentPage = 1;
            this.getView().getModel().setProperty("/results", filtered);
            this._renderPage();
        },

        _renderPage: function () {
            var data = this._filteredData;
            var totalPages = Math.ceil(data.length / this._pageSize) || 1;
            if (this._currentPage > totalPages) this._currentPage = totalPages;
            var start = (this._currentPage - 1) * this._pageSize;
            var paged = data.slice(start, start + this._pageSize);

            this.getView().getModel().setProperty("/resultsPaged", paged);

            var rowEl = this.byId("rowCount");
            if (rowEl) rowEl.setText(data.length + " rows");

            this._renderPagination(totalPages);
        },

        _renderPagination: function (totalPages) {
            var current = this._currentPage;
            if (totalPages <= 1) {
                var el = this.byId("paginationHtml");
                if (el) {
                    el.setProperty("content", '<div></div>', true);
                    var d = el.getDomRef(); if (d) d.innerHTML = '';
                }
                return;
            }

            var html = '<div class="paginationBar">';
            html += '<button class="pgBtn" onclick="window.__soApp.goToPage(1)"' + (current === 1 ? ' disabled' : '') + '>&laquo;</button>';
            html += '<button class="pgBtn" onclick="window.__soApp.goToPage(' + (current - 1) + ')"' + (current === 1 ? ' disabled' : '') + '>&lsaquo;</button>';

            var startP = Math.max(1, current - 2);
            var endP = Math.min(totalPages, current + 2);
            for (var i = startP; i <= endP; i++) {
                html += '<button class="pgBtn' + (i === current ? ' pgActive' : '') + '" onclick="window.__soApp.goToPage(' + i + ')">' + i + '</button>';
            }

            html += '<button class="pgBtn" onclick="window.__soApp.goToPage(' + (current + 1) + ')"' + (current === totalPages ? ' disabled' : '') + '>&rsaquo;</button>';
            html += '<button class="pgBtn" onclick="window.__soApp.goToPage(' + totalPages + ')"' + (current === totalPages ? ' disabled' : '') + '>&raquo;</button>';
            html += '<span class="pgInfo">Page ' + current + ' / ' + totalPages + '</span>';
            html += '</div>';

            var el = this.byId("paginationHtml");
            if (el) {
                el.setProperty("content", html, true);
                var domRef = el.getDomRef();
                if (domRef) domRef.innerHTML = html;
            }
        },

        goToPage: function (page) {
            var totalPages = Math.ceil(this._filteredData.length / this._pageSize) || 1;
            if (page < 1) page = 1;
            if (page > totalPages) page = totalPages;
            this._currentPage = page;
            this._renderPage();

            // Scroll ke atas tabel
            var scrollEl = document.getElementById("soScrollContainer");
            if (scrollEl) scrollEl.scrollTop = 0;
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
                var matchMonth = that._matchMonthFilter(r.deliveryDate, monthFilter);
                var matchReqDel = that._matchReqDeliveryFilter(r.deliveryDate);
                return matchCust && matchSearch && matchStatus && matchMonth && matchReqDel;
            });

            filtered.sort(function(a, b) {
                // Urutkan by customer dulu, lalu paling tua (daysLeft terbesar)
                var custCmp = (a.customer || "").localeCompare(b.customer || "");
                if (custCmp !== 0) return custCmp;
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

                var headers = [
                    "PO Number", "SO Number", "Customer", "Created Date",
                    "Material Code", "Material Desc",
                    "Qty Order", "Shipped", "Ots. SO", "Ots. DO", "WHFG", "Packing",
                    "Ots. SO Value", "Req. Delivery", "Days Running", "Status"
                ];
                var wsData = [headers];

                data.forEach(function (r) {
                    wsData.push([
                        r.bstnk || "",
                        r.id,
                        r.customer,
                        r.orderDateFmt,
                        r.matnr,
                        r.item,
                        Math.floor(r.kwmeng),
                        Math.floor(r.qtyGi),
                        Math.floor(r.qtyBalance2),
                        Math.floor(r.otsdo),
                        Math.floor(r.kalab),
                        Math.floor(r.kalab2),
                        r.totalValue,
                        r.deliveryDateFmt,
                        r.hariBerjalan,
                        r.status
                    ]);
                });

                var tv = data.reduce(function (s, r) { return s + r.totalValue; }, 0);
                wsData.push(["", "", "", "", "", "", "", "", "", "", "", "TOTAL", tv, "", "", ""]);

                var ws = XLSXLib.utils.aoa_to_sheet(wsData);

                // Column widths
                ws["!cols"] = [
                    { wch: 14 }, { wch: 20 }, { wch: 28 }, { wch: 14 },
                    { wch: 18 }, { wch: 32 },
                    { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
                    { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 10 }
                ];

                // Center align headers & values for Qty/Shipped/Ots/WHFG/Packing (columns G-L = index 6-11)
                var centerCols = [6, 7, 8, 9, 10, 11];
                var range = XLSXLib.utils.decode_range(ws["!ref"]);
                for (var R = range.s.r; R <= range.e.r; R++) {
                    for (var ci = 0; ci < centerCols.length; ci++) {
                        var C = centerCols[ci];
                        var addr = XLSXLib.utils.encode_cell({ r: R, c: C });
                        if (ws[addr]) {
                            if (!ws[addr].s) ws[addr].s = {};
                            ws[addr].s.alignment = { horizontal: "center" };
                        }
                    }
                }

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
                    rowPageBreak: 'avoid',
                    pageBreak: 'auto',
                    styles: { fontSize: 7, cellPadding: 1.5, overflow: 'linebreak', textColor: [0,0,0], halign: 'center', valign: 'middle' },
                    headStyles: { fillColor: [134, 179, 130], textColor: 255, fontStyle: "bold", fontSize: 7, halign: 'center', valign: 'middle' },
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
            var safeCode = (matnr || "-").replace(/</g, "&lt;");
            var rawCode  = (matnr || "-").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            var safeDesc = (maktx || "-").replace(/</g, "&lt;");

            var copyStyle = 'cursor:copy;border-radius:3px;padding:1px 2px;margin:-1px -2px;'
                          + 'transition:background 0.3s;';
            return '<div style="display:flex;flex-direction:column;line-height:1.35;">'
                 + '<span class="materialCode" onclick="event.stopPropagation();window.__soApp._copyText(\'' + rawCode + '\',this)" '
                 +   'title="Click to copy" style="' + copyStyle + '">' + safeCode + '</span>'
                 + '<span class="materialDesc">' + safeDesc + '</span>'
                 + '</div>';
        },

        formatStatus: function (status) {
            var statusClass = "";
            var statusText  = "";
            switch (status) {
                case "New": statusClass = "statusNew"; statusText = "🆕 New"; break;
                case "Pending": statusClass = "statusPending"; statusText = "Pending"; break;
                case "Low": statusClass = "statusLow"; statusText = "Low"; break;
                case "Medium": statusClass = "statusMedium"; statusText = "Medium"; break;
                case "Warning": statusClass = "statusWarning"; statusText = "⚠️ Warning"; break;
                case "Prioritas": statusClass = "statusPrioritas"; statusText = "🔥 Prioritas"; break;
                default: statusClass = "statusNew"; statusText = "🆕 New";
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

            if (lowerInput.includes("new")) filters.status = "New";
            else if (lowerInput.includes("pending")) filters.status = "Pending";
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
                if (filters.statusOutstanding && (item.status === "New" || item.status === "Pending")) return false;
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