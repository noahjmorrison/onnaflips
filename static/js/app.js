// Common utilities for Onna's Flips app

function formatCurrency(amount) {
    if (amount === null || amount === undefined) return '-';
    return '$' + Number(amount).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function formatDate(dateStr) {
    if (!dateStr) return '-';
    const d = new Date(dateStr + 'T00:00:00');
    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + (d.getFullYear() % 100);
}

function formatPercent(decimal) {
    if (decimal === null || decimal === undefined) return '-';
    return Math.round(decimal * 100) + '%';
}
