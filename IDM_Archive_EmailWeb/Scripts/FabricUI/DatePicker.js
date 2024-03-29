﻿
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

(function ($) {

    /**
     * DatePicker Plugin
     */

    $.fn.DatePicker = function (options) {

        return this.each(function () {

            /** Set up variables and run the Pickadate plugin. */
            var $datePicker = $(this);
            var $dateField = $datePicker.find('.ms-TextField-field').pickadate($.extend({
                // Strings and translations.
                weekdaysShort: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

                // Don't render the buttons
                today: '',
                clear: '',
                close: '',

                // Events
                onStart: function () {
                    initCustomView($datePicker);
                },

                // Classes
                klass: {

                    // The element states
                    input: 'ms-DatePicker-input',
                    active: 'ms-DatePicker-input--active',

                    // The root picker and states
                    picker: 'ms-DatePicker-picker',
                    opened: 'ms-DatePicker-picker--opened',
                    focused: 'ms-DatePicker-picker--focused',

                    // The picker holder
                    holder: 'ms-DatePicker-holder',

                    // The picker frame, wrapper, and box
                    frame: 'ms-DatePicker-frame',
                    wrap: 'ms-DatePicker-wrap',
                    box: 'ms-DatePicker-dayPicker',

                    // The picker header
                    header: 'ms-DatePicker-header',

                    // Month & year labels
                    month: 'ms-DatePicker-month',
                    year: 'ms-DatePicker-year',

                    // Table of dates
                    table: 'ms-DatePicker-table',

                    // Weekday labels
                    weekdays: 'ms-DatePicker-weekday',

                    // Day states
                    day: 'ms-DatePicker-day',
                    disabled: 'ms-DatePicker-day--disabled',
                    selected: 'ms-DatePicker-day--selected',
                    highlighted: 'ms-DatePicker-day--highlighted',
                    now: 'ms-DatePicker-day--today',
                    infocus: 'ms-DatePicker-day--infocus',
                    outfocus: 'ms-DatePicker-day--outfocus',

                }
            }, options || {}));
            var $picker = $dateField.pickadate('picker');

            /** Respond to built-in picker events. */
            $picker.on({
                render: function () {
                    updateCustomView($datePicker);
                },
                open: function () {
                    scrollUp($datePicker);
                }
            });

        });
    };

    /**
     * After the Pickadate plugin starts, this function
     * adds additional controls to the picker view.
     */
    function initCustomView($datePicker) {

        /** Get some variables ready. */
        var $monthControls = $datePicker.find('.ms-DatePicker-monthComponents');
        var $goToday = $datePicker.find('.ms-DatePicker-goToday');
        var $monthPicker = $datePicker.find('.ms-DatePicker-monthPicker');
        var $yearPicker = $datePicker.find('.ms-DatePicker-yearPicker');
        var $pickerWrapper = $datePicker.find('.ms-DatePicker-wrap');
        var $picker = $datePicker.find('.ms-TextField-field').pickadate('picker');

        /** Move the month picker into position. */
        $monthControls.appendTo($pickerWrapper);
        $goToday.appendTo($pickerWrapper);
        $monthPicker.appendTo($pickerWrapper);
        $yearPicker.appendTo($pickerWrapper);

        /** Update the custom view. */
        updateCustomView($datePicker);

        /** Move back one month. */
        $monthControls.on('click', '.js-prevMonth', function (event) {
            event.preventDefault();
            var newMonth = $picker.get('highlight').month - 1;
            changeHighlightedDate($picker, null, newMonth, null);
        });

        /** Move ahead one month. */
        $monthControls.on('click', '.js-nextMonth', function (event) {
            event.preventDefault();
            var newMonth = $picker.get('highlight').month + 1;
            changeHighlightedDate($picker, null, newMonth, null);
        });

        /** Move back one year. */
        $monthPicker.on('click', '.js-prevYear', function (event) {
            event.preventDefault();
            var newYear = $picker.get('highlight').year - 1;
            changeHighlightedDate($picker, newYear, null, null);
        });

        /** Move ahead one year. */
        $monthPicker.on('click', '.js-nextYear', function (event) {
            event.preventDefault();
            var newYear = $picker.get('highlight').year + 1;
            changeHighlightedDate($picker, newYear, null, null);
        });

        /** Move back one decade. */
        $yearPicker.on('click', '.js-prevDecade', function (event) {
            event.preventDefault();
            var newYear = $picker.get('highlight').year - 10;
            changeHighlightedDate($picker, newYear, null, null);
        });

        /** Move ahead one decade. */
        $yearPicker.on('click', '.js-nextDecade', function (event) {
            event.preventDefault();
            var newYear = $picker.get('highlight').year + 10;
            changeHighlightedDate($picker, newYear, null, null);
        });

        /** Go to the current date, shown in the day picking view. */
        $goToday.click(function (event) {
            event.preventDefault();

            /** Select the current date, while keeping the picker open. */
            var now = new Date();
            $picker.set('select', [now.getFullYear(), now.getMonth(), now.getDate()]);

            /** Switch to the default (calendar) view. */
            $datePicker.removeClass('is-pickingMonths').removeClass('is-pickingYears');

        });

        /** Change the highlighted month. */
        $monthPicker.on('click', '.js-changeDate', function (event) {
            event.preventDefault();

            /** Get the requested date from the data attributes. */
            var newYear = $(this).attr('data-year');
            var newMonth = $(this).attr('data-month');
            var newDay = $(this).attr('data-day');

            /** Update the date. */
            changeHighlightedDate($picker, newYear, newMonth, newDay);

            /** If we've been in the "picking months" state on mobile, remove that state so we show the calendar again. */
            if ($datePicker.hasClass('is-pickingMonths')) {
                $datePicker.removeClass('is-pickingMonths');
            }
        });

        /** Change the highlighted year. */
        $yearPicker.on('click', '.js-changeDate', function (event) {
            event.preventDefault();

            /** Get the requested date from the data attributes. */
            var newYear = $(this).attr('data-year');
            var newMonth = $(this).attr('data-month');
            var newDay = $(this).attr('data-day');

            /** Update the date. */
            changeHighlightedDate($picker, newYear, newMonth, newDay);

            /** If we've been in the "picking years" state on mobile, remove that state so we show the calendar again. */
            if ($datePicker.hasClass('is-pickingYears')) {
                $datePicker.removeClass('is-pickingYears');
            }
        });

        /** Switch to the default state. */
        $monthPicker.on('click', '.js-showDayPicker', function () {
            $datePicker.removeClass('is-pickingMonths');
            $datePicker.removeClass('is-pickingYears');
        });

        /** Switch to the is-pickingMonths state. */
        $monthControls.on('click', '.js-showMonthPicker', function () {
            $datePicker.toggleClass('is-pickingMonths');
        });

        /** Switch to the is-pickingYears state. */
        $monthPicker.on('click', '.js-showYearPicker', function () {
            $datePicker.toggleClass('is-pickingYears');
        });

    }

    /** Change the highlighted date. */
    function changeHighlightedDate($picker, newYear, newMonth, newDay) {

        /** All letiables are optional. If not provided, default to the current value. */
        if (typeof newYear === "" || newYear === null || typeof newYear === "undefined") {
            newYear = $picker.get("highlight").year;
        }
        if (typeof newMonth === "" || newMonth === null || typeof newMonth === "undefined") {
            newMonth = $picker.get("highlight").month;
        }
        if (typeof newDay === "" || newDay === null || typeof newDay === "undefined") {
            newDay = $picker.get("highlight").date;
        }

        /** Update it. */
        $picker.set('highlight', [newYear, newMonth, newDay]);

    }


    /** Whenever the picker renders, do our own rendering on the custom controls. */
    function updateCustomView($datePicker) {

        /** Get some variables ready. */
        var $monthPicker = $datePicker.find('.ms-DatePicker-monthPicker');
        var $yearPicker = $datePicker.find('.ms-DatePicker-yearPicker');
        var $picker = $datePicker.find('.ms-TextField-field').pickadate('picker');

        /** Set the correct year. */
        $monthPicker.find('.ms-DatePicker-currentYear').text($picker.get('view').year);

        /** Highlight the current month. */
        $monthPicker.find('.ms-DatePicker-monthOption').removeClass('is-highlighted');
        $monthPicker.find('.ms-DatePicker-monthOption[data-month="' + $picker.get('highlight').month + '"]').addClass('is-highlighted');

        /** Generate the grid of years for the year picker view. */

        // Start by removing any existing generated output. */
        $yearPicker.find('.ms-DatePicker-currentDecade').remove();
        $yearPicker.find('.ms-DatePicker-optionGrid').remove();

        // Generate the output by going through the years.
        var startingYear = $picker.get('highlight').year - 11;
        var decadeText = startingYear + " - " + (startingYear + 11);
        var output = '<div class="ms-DatePicker-currentDecade">' + decadeText + '</div>';
        output += '<div class="ms-DatePicker-optionGrid">';
        for (var year = startingYear; year < (startingYear + 12) ; year++) {
            output += '<span class="ms-DatePicker-yearOption js-changeDate" data-year="' + year + '">' + year + '</span>';
        }
        output += '</div>';

        // Output the title and grid of years generated above.
        $yearPicker.append(output);

        /** Highlight the current year. */
        $yearPicker.find('.ms-DatePicker-yearOption').removeClass('is-highlighted');
        $yearPicker.find('.ms-DatePicker-yearOption[data-year="' + $picker.get('highlight').year + '"]').addClass('is-highlighted');
    }

    /** Scroll the page up so that the field the date picker is attached to is at the top. */
    function scrollUp($datePicker) {
        $('html, body').animate({
            scrollTop: $datePicker.offset().top
        }, 367);
    }

})(jQuery);
!function (a) { "function" == typeof define && define.amd ? define("picker", ["jquery"], a) : "object" == typeof exports ? module.exports = a(require("jquery")) : this.Picker = a(jQuery) }(function (a) { function b(f, g, h, k) { function l() { return b._.node("div", b._.node("div", b._.node("div", b._.node("div", w.component.nodes(r.open), t.box), t.wrap), t.frame), t.holder) } function m() { u.data(g, w).addClass(t.input).val(u.data("value") ? w.get("select", s.format) : f.value).on("focus." + r.id + " click." + r.id, p), s.editable || u.on("keydown." + r.id, function (a) { var b = a.keyCode, c = /^(8|46)$/.test(b); return 27 == b ? (w.close(), !1) : void ((32 == b || c || !r.open && w.component.key[b]) && (a.preventDefault(), a.stopPropagation(), c ? w.clear().close() : w.open())) }), e(f, { haspopup: !0, expanded: !1, readonly: !1, owns: f.id + "_root" + (w._hidden ? " " + w._hidden.id : "") }) } function n() { w.$root.on({ focusin: function (a) { w.$root.removeClass(t.focused), a.stopPropagation() }, "mousedown click": function (b) { var c = b.target; c != w.$root.children()[0] && (b.stopPropagation(), "mousedown" != b.type || a(c).is(":input") || "OPTION" == c.nodeName || (b.preventDefault(), f.focus())) } }).on("click", "[data-pick], [data-nav], [data-clear], [data-close]", function () { var b = a(this), c = b.data(), d = b.hasClass(t.navDisabled) || b.hasClass(t.disabled), e = document.activeElement; e = e && (e.type || e.href) && e, (d || e && !a.contains(w.$root[0], e)) && f.focus(), !d && c.nav ? w.set("highlight", w.component.item.highlight, { nav: c.nav }) : !d && "pick" in c ? w.set("select", c.pick).close(!0) : c.clear ? w.clear().close(!0) : c.close && w.close(!0) }), e(w.$root[0], "hidden", !0) } function o() { var b; s.hiddenName === !0 ? (b = f.name, f.name = "") : (b = ["string" == typeof s.hiddenPrefix ? s.hiddenPrefix : "", "string" == typeof s.hiddenSuffix ? s.hiddenSuffix : "_submit"], b = b[0] + f.name + b[1]), w._hidden = a('<input type=hidden name="' + b + '"' + (u.data("value") || f.value ? ' value="' + w.get("select", s.formatSubmit) + '"' : "") + ">")[0], u.on("change." + r.id, function () { w._hidden.value = f.value ? w.get("select", s.formatSubmit) : "" }).after(w._hidden) } function p(a) { a.stopPropagation(), "focus" == a.type && w.$root.addClass(t.focused), w.open() } if (!f) return b; var q = !1, r = { id: f.id || "P" + Math.abs(~~(Math.random() * new Date)) }, s = h ? a.extend(!0, {}, h.defaults, k) : k || {}, t = a.extend({}, b.klasses(), s.klass), u = a(f), v = function () { return this.start() }, w = v.prototype = { constructor: v, $node: u, start: function () { return r && r.start ? w : (r.methods = {}, r.start = !0, r.open = !1, r.type = f.type, f.autofocus = f == document.activeElement, f.readOnly = !s.editable, f.id = f.id || r.id, "text" != f.type && (f.type = "text"), w.component = new h(w, s), w.$root = a(b._.node("div", l(), t.picker, 'id="' + f.id + '_root"')), n(), s.formatSubmit && o(), m(), s.container ? a(s.container).append(w.$root) : u.after(w.$root), w.on({ start: w.component.onStart, render: w.component.onRender, stop: w.component.onStop, open: w.component.onOpen, close: w.component.onClose, set: w.component.onSet }).on({ start: s.onStart, render: s.onRender, stop: s.onStop, open: s.onOpen, close: s.onClose, set: s.onSet }), q = c(w.$root.children()[0]), f.autofocus && w.open(), w.trigger("start").trigger("render")) }, render: function (a) { return a ? w.$root.html(l()) : w.$root.find("." + t.box).html(w.component.nodes(r.open)), w.trigger("render") }, stop: function () { return r.start ? (w.close(), w._hidden && w._hidden.parentNode.removeChild(w._hidden), w.$root.remove(), u.removeClass(t.input).removeData(g), setTimeout(function () { u.off("." + r.id) }, 0), f.type = r.type, f.readOnly = !1, w.trigger("stop"), r.methods = {}, r.start = !1, w) : w }, open: function (c) { return r.open ? w : (u.addClass(t.active), e(f, "expanded", !0), setTimeout(function () { w.$root.addClass(t.opened), e(w.$root[0], "hidden", !1) }, 0), c !== !1 && (r.open = !0, q && j.css("overflow", "hidden").css("padding-right", "+=" + d()), u.trigger("focus"), i.on("click." + r.id + " focusin." + r.id, function (a) { var b = a.target; b != f && b != document && 3 != a.which && w.close(b === w.$root.children()[0]) }).on("keydown." + r.id, function (c) { var d = c.keyCode, e = w.component.key[d], g = c.target; 27 == d ? w.close(!0) : g != f || !e && 13 != d ? a.contains(w.$root[0], g) && 13 == d && (c.preventDefault(), g.click()) : (c.preventDefault(), e ? b._.trigger(w.component.key.go, w, [b._.trigger(e)]) : w.$root.find("." + t.highlighted).hasClass(t.disabled) || w.set("select", w.component.item.highlight).close()) })), w.trigger("open")) }, close: function (a) { return a && (u.off("focus." + r.id).trigger("focus"), setTimeout(function () { u.on("focus." + r.id, p) }, 0)), u.removeClass(t.active), e(f, "expanded", !1), setTimeout(function () { w.$root.removeClass(t.opened + " " + t.focused), e(w.$root[0], "hidden", !0) }, 0), r.open ? (r.open = !1, q && j.css("overflow", "").css("padding-right", "-=" + d()), i.off("." + r.id), w.trigger("close")) : w }, clear: function (a) { return w.set("clear", null, a) }, set: function (b, c, d) { var e, f, g = a.isPlainObject(b), h = g ? b : {}; if (d = g && a.isPlainObject(c) ? c : d || {}, b) { g || (h[b] = c); for (e in h) f = h[e], e in w.component.item && (void 0 === f && (f = null), w.component.set(e, f, d)), ("select" == e || "clear" == e) && u.val("clear" == e ? "" : w.get(e, s.format)).trigger("change"); w.render() } return d.muted ? w : w.trigger("set", h) }, get: function (a, c) { if (a = a || "value", null != r[a]) return r[a]; if ("value" == a) return f.value; if (a in w.component.item) { if ("string" == typeof c) { var d = w.component.get(a); return d ? b._.trigger(w.component.formats.toString, w.component, [c, d]) : "" } return w.component.get(a) } }, on: function (b, c, d) { var e, f, g = a.isPlainObject(b), h = g ? b : {}; if (b) { g || (h[b] = c); for (e in h) f = h[e], d && (e = "_" + e), r.methods[e] = r.methods[e] || [], r.methods[e].push(f) } return w }, off: function () { var a, b, c = arguments; for (a = 0, namesCount = c.length; namesCount > a; a += 1) b = c[a], b in r.methods && delete r.methods[b]; return w }, trigger: function (a, c) { var d = function (a) { var d = r.methods[a]; d && d.map(function (a) { b._.trigger(a, w, [c]) }) }; return d("_" + a), d(a), w } }; return new v } function c(a) { var b, c = "position"; return a.currentStyle ? b = a.currentStyle[c] : window.getComputedStyle && (b = getComputedStyle(a)[c]), "fixed" == b } function d() { if (j.height() <= h.height()) return 0; var b = a('<div style="visibility:hidden;width:100px" />').appendTo("body"), c = b[0].offsetWidth; b.css("overflow", "scroll"); var d = a('<div style="width:100%" />').appendTo(b), e = d[0].offsetWidth; return b.remove(), c - e } function e(b, c, d) { if (a.isPlainObject(c)) for (var e in c) f(b, e, c[e]); else f(b, c, d) } function f(a, b, c) { a.setAttribute(("role" == b ? "" : "aria-") + b, c) } function g(b, c) { a.isPlainObject(b) || (b = { attribute: c }), c = ""; for (var d in b) { var e = ("role" == d ? "" : "aria-") + d, f = b[d]; c += null == f ? "" : e + '="' + b[d] + '"' } return c } var h = a(window), i = a(document), j = a(document.documentElement); return b.klasses = function (a) { return a = a || "picker", { picker: a, opened: a + "--opened", focused: a + "--focused", input: a + "__input", active: a + "__input--active", holder: a + "__holder", frame: a + "__frame", wrap: a + "__wrap", box: a + "__box" } }, b._ = { group: function (a) { for (var c, d = "", e = b._.trigger(a.min, a) ; e <= b._.trigger(a.max, a, [e]) ; e += a.i) c = b._.trigger(a.item, a, [e]), d += b._.node(a.node, c[0], c[1], c[2]); return d }, node: function (b, c, d, e) { return c ? (c = a.isArray(c) ? c.join("") : c, d = d ? ' class="' + d + '"' : "", e = e ? " " + e : "", "<" + b + d + e + ">" + c + "</" + b + ">") : "" }, lead: function (a) { return (10 > a ? "0" : "") + a }, trigger: function (a, b, c) { return "function" == typeof a ? a.apply(b, c || []) : a }, digits: function (a) { return /\d/.test(a[1]) ? 2 : 1 }, isDate: function (a) { return {}.toString.call(a).indexOf("Date") > -1 && this.isInteger(a.getUTCDate()) }, isInteger: function (a) { return {}.toString.call(a).indexOf("Number") > -1 && a % 1 === 0 }, ariaAttr: g }, b.extend = function (c, d) { a.fn[c] = function (e, f) { var g = this.data(c); return "picker" == e ? g : g && "string" == typeof e ? b._.trigger(g[e], g, [f]) : this.each(function () { var f = a(this); f.data(c) || new b(this, c, d, e) }) }, a.fn[c].defaults = d.defaults }, b }), function (a) { "function" == typeof define && define.amd ? define(["picker", "jquery"], a) : "object" == typeof exports ? module.exports = a(require("./picker.js"), require("jquery")) : a(Picker, jQuery) }(function (a, b) { function c(a, b) { var c = this, d = a.$node[0], e = d.value, f = a.$node.data("value"), g = f || e, h = f ? b.formatSubmit : b.format, i = function () { return d.currentStyle ? "rtl" == d.currentStyle.direction : "rtl" == getComputedStyle(a.$root[0]).direction }; c.settings = b, c.$node = a.$node, c.queue = { min: "measure create", max: "measure create", now: "now create", select: "parse create validate", highlight: "parse navigate create validate", view: "parse create validate viewset", disable: "deactivate", enable: "activate" }, c.item = {}, c.item.clear = null, c.item.disable = (b.disable || []).slice(0), c.item.enable = -function (a) { return a[0] === !0 ? a.shift() : -1 }(c.item.disable), c.set("min", b.min).set("max", b.max).set("now"), g ? c.set("select", g, { format: h }) : c.set("select", null).set("highlight", c.item.now), c.key = { 40: 7, 38: -7, 39: function () { return i() ? -1 : 1 }, 37: function () { return i() ? 1 : -1 }, go: function (a) { var b = c.item.highlight, d = new Date(Date.UTC(b.year, b.month, b.date + a)); c.set("highlight", d, { interval: a }), this.render() } }, a.on("render", function () { a.$root.find("." + b.klass.selectMonth).on("change", function () { var c = this.value; c && (a.set("highlight", [a.get("view").year, c, a.get("highlight").date]), a.$root.find("." + b.klass.selectMonth).trigger("focus")) }), a.$root.find("." + b.klass.selectYear).on("change", function () { var c = this.value; c && (a.set("highlight", [c, a.get("view").month, a.get("highlight").date]), a.$root.find("." + b.klass.selectYear).trigger("focus")) }) }, 1).on("open", function () { var d = ""; c.disabled(c.get("now")) && (d = ":not(." + b.klass.buttonToday + ")"), a.$root.find("button" + d + ", select").attr("disabled", !1) }, 1).on("close", function () { a.$root.find("button, select").attr("disabled", !0) }, 1) } var d = 7, e = 6, f = a._; c.prototype.set = function (a, b, c) { var d = this, e = d.item; return null === b ? ("clear" == a && (a = "select"), e[a] = b, d) : (e["enable" == a ? "disable" : "flip" == a ? "enable" : a] = d.queue[a].split(" ").map(function (e) { return b = d[e](a, b, c) }).pop(), "select" == a ? d.set("highlight", e.select, c) : "highlight" == a ? d.set("view", e.highlight, c) : a.match(/^(flip|min|max|disable|enable)$/) && (e.select && d.disabled(e.select) && d.set("select", e.select, c), e.highlight && d.disabled(e.highlight) && d.set("highlight", e.highlight, c)), d) }, c.prototype.get = function (a) { return this.item[a] }, c.prototype.create = function (a, c, d) { var e, g = this; return c = void 0 === c ? a : c, c == -1 / 0 || 1 / 0 == c ? e = c : b.isPlainObject(c) && f.isInteger(c.pick) ? c = c.obj : b.isArray(c) ? (c = new Date(Date.UTC(c[0], c[1], c[2])), c = f.isDate(c) ? c : g.create().obj) : c = f.isInteger(c) ? g.normalize(new Date(c), d) : f.isDate(c) ? g.normalize(c, d) : g.now(a, c, d), { year: e || c.getUTCFullYear(), month: e || c.getUTCMonth(), date: e || c.getUTCDate(), day: e || c.getUTCDay(), obj: e || c, pick: e || c.getTime() } }, c.prototype.createRange = function (a, c) { var d = this, e = function (a) { return a === !0 || b.isArray(a) || f.isDate(a) ? d.create(a) : a }; return f.isInteger(a) || (a = e(a)), f.isInteger(c) || (c = e(c)), f.isInteger(a) && b.isPlainObject(c) ? a = [c.year, c.month, c.date + a] : f.isInteger(c) && b.isPlainObject(a) && (c = [a.year, a.month, a.date + c]), { from: e(a), to: e(c) } }, c.prototype.withinRange = function (a, b) { return a = this.createRange(a.from, a.to), b.pick >= a.from.pick && b.pick <= a.to.pick }, c.prototype.overlapRanges = function (a, b) { var c = this; return a = c.createRange(a.from, a.to), b = c.createRange(b.from, b.to), c.withinRange(a, b.from) || c.withinRange(a, b.to) || c.withinRange(b, a.from) || c.withinRange(b, a.to) }, c.prototype.now = function (a, b, c) { return b = new Date, c && c.rel && b.setUTCDate(b.getUTCDate() + c.rel), this.normalize(b, c) }, c.prototype.navigate = function (a, c, d) { var e, f, g, h, i = b.isArray(c), j = b.isPlainObject(c), k = this.item.view; if (i || j) { for (j ? (f = c.year, g = c.month, h = c.date) : (f = +c[0], g = +c[1], h = +c[2]), d && d.nav && k && k.month !== g && (f = k.year, g = k.month), e = new Date(Date.UTC(f, g + (d && d.nav ? d.nav : 0), 1)), f = e.getUTCFullYear(), g = e.getUTCMonth() ; new Date(Date.UTC(f, g, h)).getUTCMonth() !== g;) h -= 1; c = [f, g, h] } return c }, c.prototype.normalize = function (a) { return a.setUTCHours(0, 0, 0, 0), a }, c.prototype.measure = function (a, b) { var c = this; return b ? "string" == typeof b ? b = c.parse(a, b) : f.isInteger(b) && (b = c.now(a, b, { rel: b })) : b = "min" == a ? -1 / 0 : 1 / 0, b }, c.prototype.viewset = function (a, b) { return this.create([b.year, b.month, 1]) }, c.prototype.validate = function (a, c, d) { var e, g, h, i, j = this, k = c, l = d && d.interval ? d.interval : 1, m = -1 === j.item.enable, n = j.item.min, o = j.item.max, p = m && j.item.disable.filter(function (a) { if (b.isArray(a)) { var d = j.create(a).pick; d < c.pick ? e = !0 : d > c.pick && (g = !0) } return f.isInteger(a) }).length; if ((!d || !d.nav) && (!m && j.disabled(c) || m && j.disabled(c) && (p || e || g) || !m && (c.pick <= n.pick || c.pick >= o.pick))) for (m && !p && (!g && l > 0 || !e && 0 > l) && (l *= -1) ; j.disabled(c) && (Math.abs(l) > 1 && (c.month < k.month || c.month > k.month) && (c = k, l = l > 0 ? 1 : -1), c.pick <= n.pick ? (h = !0, l = 1, c = j.create([n.year, n.month, n.date + (c.pick === n.pick ? 0 : -1)])) : c.pick >= o.pick && (i = !0, l = -1, c = j.create([o.year, o.month, o.date + (c.pick === o.pick ? 0 : 1)])), !h || !i) ;) c = j.create([c.year, c.month, c.date + l]); return c }, c.prototype.disabled = function (a) { var c = this, d = c.item.disable.filter(function (d) { return f.isInteger(d) ? a.day === (c.settings.firstDay ? d : d - 1) % 7 : b.isArray(d) || f.isDate(d) ? a.pick === c.create(d).pick : b.isPlainObject(d) ? c.withinRange(d, a) : void 0 }); return d = d.length && !d.filter(function (a) { return b.isArray(a) && "inverted" == a[3] || b.isPlainObject(a) && a.inverted }).length, -1 === c.item.enable ? !d : d || a.pick < c.item.min.pick || a.pick > c.item.max.pick }, c.prototype.parse = function (a, b, c) { var d = this, e = {}; return b && "string" == typeof b ? (c && c.format || (c = c || {}, c.format = d.settings.format), d.formats.toArray(c.format).map(function (a) { var c = d.formats[a], g = c ? f.trigger(c, d, [b, e]) : a.replace(/^!/, "").length; c && (e[a] = b.substr(0, g)), b = b.substr(g) }), [e.yyyy || e.yy, +(e.mm || e.m) - 1, e.dd || e.d]) : b }, c.prototype.formats = function () { function a(a, b, c) { var d = a.match(/\w+/)[0]; return c.mm || c.m || (c.m = b.indexOf(d) + 1), d.length } function b(a) { return a.match(/\w+/)[0].length } return { d: function (a, b) { return a ? f.digits(a) : b.date }, dd: function (a, b) { return a ? 2 : f.lead(b.date) }, ddd: function (a, c) { return a ? b(a) : this.settings.weekdaysShort[c.day] }, dddd: function (a, c) { return a ? b(a) : this.settings.weekdaysFull[c.day] }, m: function (a, b) { return a ? f.digits(a) : b.month + 1 }, mm: function (a, b) { return a ? 2 : f.lead(b.month + 1) }, mmm: function (b, c) { var d = this.settings.monthsShort; return b ? a(b, d, c) : d[c.month] }, mmmm: function (b, c) { var d = this.settings.monthsFull; return b ? a(b, d, c) : d[c.month] }, yy: function (a, b) { return a ? 2 : ("" + b.year).slice(2) }, yyyy: function (a, b) { return a ? 4 : b.year }, toArray: function (a) { return a.split(/(d{1,4}|m{1,4}|y{4}|yy|!.)/g) }, toString: function (a, b) { var c = this; return c.formats.toArray(a).map(function (a) { return f.trigger(c.formats[a], c, [0, b]) || a.replace(/^!/, "") }).join("") } } }(), c.prototype.isDateExact = function (a, c) { var d = this; return f.isInteger(a) && f.isInteger(c) || "boolean" == typeof a && "boolean" == typeof c ? a === c : (f.isDate(a) || b.isArray(a)) && (f.isDate(c) || b.isArray(c)) ? d.create(a).pick === d.create(c).pick : b.isPlainObject(a) && b.isPlainObject(c) ? d.isDateExact(a.from, c.from) && d.isDateExact(a.to, c.to) : !1 }, c.prototype.isDateOverlap = function (a, c) { var d = this, e = d.settings.firstDay ? 1 : 0; return f.isInteger(a) && (f.isDate(c) || b.isArray(c)) ? (a = a % 7 + e, a === d.create(c).day + 1) : f.isInteger(c) && (f.isDate(a) || b.isArray(a)) ? (c = c % 7 + e, c === d.create(a).day + 1) : b.isPlainObject(a) && b.isPlainObject(c) ? d.overlapRanges(a, c) : !1 }, c.prototype.flipEnable = function (a) { var b = this.item; b.enable = a || (-1 == b.enable ? 1 : -1) }, c.prototype.deactivate = function (a, c) { var d = this, e = d.item.disable.slice(0); return "flip" == c ? d.flipEnable() : c === !1 ? (d.flipEnable(1), e = []) : c === !0 ? (d.flipEnable(-1), e = []) : c.map(function (a) { for (var c, g = 0; g < e.length; g += 1) if (d.isDateExact(a, e[g])) { c = !0; break } c || (f.isInteger(a) || f.isDate(a) || b.isArray(a) || b.isPlainObject(a) && a.from && a.to) && e.push(a) }), e }, c.prototype.activate = function (a, c) { var d = this, e = d.item.disable, g = e.length; return "flip" == c ? d.flipEnable() : c === !0 ? (d.flipEnable(1), e = []) : c === !1 ? (d.flipEnable(-1), e = []) : c.map(function (a) { var c, h, i, j; for (i = 0; g > i; i += 1) { if (h = e[i], d.isDateExact(h, a)) { c = e[i] = null, j = !0; break } if (d.isDateOverlap(h, a)) { b.isPlainObject(a) ? (a.inverted = !0, c = a) : b.isArray(a) ? (c = a, c[3] || c.push("inverted")) : f.isDate(a) && (c = [a.getUTCFullYear(), a.getUTCMonth(), a.getUTCDate(), "inverted"]); break } } if (c) for (i = 0; g > i; i += 1) if (d.isDateExact(e[i], a)) { e[i] = null; break } if (j) for (i = 0; g > i; i += 1) if (d.isDateOverlap(e[i], a)) { e[i] = null; break } c && e.push(c) }), e.filter(function (a) { return null != a }) }, c.prototype.nodes = function (a) { var b = this, c = b.settings, g = b.item, h = g.now, i = g.select, j = g.highlight, k = g.view, l = g.disable, m = g.min, n = g.max, o = function (a, b) { return c.firstDay && (a.push(a.shift()), b.push(b.shift())), f.node("thead", f.node("tr", f.group({ min: 0, max: d - 1, i: 1, node: "th", item: function (d) { return [a[d], c.klass.weekdays, 'scope=col title="' + b[d] + '"'] } }))) }((c.showWeekdaysFull ? c.weekdaysFull : c.weekdaysShort).slice(0), c.weekdaysFull.slice(0)), p = function (a) { return f.node("div", " ", c.klass["nav" + (a ? "Next" : "Prev")] + (a && k.year >= n.year && k.month >= n.month || !a && k.year <= m.year && k.month <= m.month ? " " + c.klass.navDisabled : ""), "data-nav=" + (a || -1) + " " + f.ariaAttr({ role: "button", components: b.$node[0].id + "_table" }) + ' title="' + (a ? c.labelMonthNext : c.labelMonthPrev) + '"') }, q = function () { var d = c.showMonthsShort ? c.monthsShort : c.monthsFull; return c.selectMonths ? f.node("select", f.group({ min: 0, max: 11, i: 1, node: "option", item: function (a) { return [d[a], 0, "value=" + a + (k.month == a ? " selected" : "") + (k.year == m.year && a < m.month || k.year == n.year && a > n.month ? " disabled" : "")] } }), c.klass.selectMonth, (a ? "" : "disabled") + " " + f.ariaAttr({ components: b.$node[0].id + "_table" }) + ' title="' + c.labelMonthSelect + '"') : f.node("div", d[k.month], c.klass.month) }, r = function () { var d = k.year, e = c.selectYears === !0 ? 5 : ~~(c.selectYears / 2); if (e) { var g = m.year, h = n.year, i = d - e, j = d + e; if (g > i && (j += g - i, i = g), j > h) { var l = i - g, o = j - h; i -= l > o ? o : l, j = h } return f.node("select", f.group({ min: i, max: j, i: 1, node: "option", item: function (a) { return [a, 0, "value=" + a + (d == a ? " selected" : "")] } }), c.klass.selectYear, (a ? "" : "disabled") + " " + f.ariaAttr({ components: b.$node[0].id + "_table" }) + ' title="' + c.labelYearSelect + '"') } return f.node("div", d, c.klass.year) }; return f.node("div", (c.selectYears ? r() + q() : q() + r()) + p() + p(1), c.klass.header) + f.node("table", o + f.node("tbody", f.group({ min: 0, max: e - 1, i: 1, node: "tr", item: function (a) { var e = c.firstDay && 0 === b.create([k.year, k.month, 1]).day ? -7 : 0; return [f.group({ min: d * a - k.day + e + 1, max: function () { return this.min + d - 1 }, i: 1, node: "td", item: function (a) { a = b.create([k.year, k.month, a + (c.firstDay ? 1 : 0)]); var d = i && i.pick == a.pick, e = j && j.pick == a.pick, g = l && b.disabled(a) || a.pick < m.pick || a.pick > n.pick; return [f.node("div", a.date, function (b) { return b.push(k.month == a.month ? c.klass.infocus : c.klass.outfocus), h.pick == a.pick && b.push(c.klass.now), d && b.push(c.klass.selected), e && b.push(c.klass.highlighted), g && b.push(c.klass.disabled), b.join(" ") }([c.klass.day]), "data-pick=" + a.pick + " " + f.ariaAttr({ role: "gridcell", selected: d && b.$node.val() === f.trigger(b.formats.toString, b, [c.format, a]) ? !0 : null, activedescendant: e ? !0 : null, disabled: g ? !0 : null })), "", f.ariaAttr({ role: "presentation" })] } })] } })), c.klass.table, 'id="' + b.$node[0].id + '_table" ' + f.ariaAttr({ role: "grid", components: b.$node[0].id, readonly: !0 })) + f.node("div", f.node("button", c.today, c.klass.buttonToday, "type=button data-pick=" + h.pick + (a && !b.disabled(h) ? "" : " disabled") + " " + f.ariaAttr({ components: b.$node[0].id })) + f.node("button", c.clear, c.klass.buttonClear, "type=button data-clear=1" + (a ? "" : " disabled") + " " + f.ariaAttr({ components: b.$node[0].id })) + f.node("button", c.close, c.klass.buttonClose, "type=button data-close=true " + (a ? "" : " disabled") + " " + f.ariaAttr({ components: b.$node[0].id })), c.klass.footer) }, c.defaults = function (a) { return { labelMonthNext: "Next month", labelMonthPrev: "Previous month", labelMonthSelect: "Select a month", labelYearSelect: "Select a year", monthsFull: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], weekdaysFull: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], weekdaysShort: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], today: "Today", clear: "Clear", close: "Close", format: "d mmmm, yyyy", klass: { table: a + "table", header: a + "header", navPrev: a + "nav--prev", navNext: a + "nav--next", navDisabled: a + "nav--disabled", month: a + "month", year: a + "year", selectMonth: a + "select--month", selectYear: a + "select--year", weekdays: a + "weekday", day: a + "day", disabled: a + "day--disabled", selected: a + "day--selected", highlighted: a + "day--highlighted", now: a + "day--today", infocus: a + "day--infocus", outfocus: a + "day--outfocus", footer: a + "footer", buttonClear: a + "button--clear", buttonToday: a + "button--today", buttonClose: a + "button--close" } } }(a.klasses().picker + "__"), a.extend("pickadate", c) });