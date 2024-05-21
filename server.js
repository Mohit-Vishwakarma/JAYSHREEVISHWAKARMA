const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');

const app = express();
const port = 3000;

// Middleware
app.use(bodyParser.json());

// Load or create Excel file
const excelFilePath = 'orders.xlsx';

let workbook;
if (fs.existsSync(excelFilePath)) {
    workbook = xlsx.readFile(excelFilePath);
} else {
    workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet([[
        'id', 'name', 'address', 'contact', 'dateOfCreation', 'orderDetails', 
        'totalAmount', 'advanceAmount', 'challanDetail', 'orderCompletionStatus'
    ]]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Orders');
    xlsx.writeFile(workbook, excelFilePath);
}

const getOrdersSheet = () => {
    return workbook.Sheets['Orders'];
};

const saveWorkbook = () => {
    xlsx.writeFile(workbook, excelFilePath);
};

const readOrders = () => {
    const worksheet = getOrdersSheet();
    return xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);
};

const appendOrder = (order) => {
    const worksheet = getOrdersSheet();
    const orders = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    orders.push([
        order.id, order.name, order.address, order.contact,
        order.dateOfCreation, order.orderDetails, order.totalAmount,
        order.advanceAmount, order.challanDetail, order.orderCompletionStatus
    ]);
    const newWorksheet = xlsx.utils.aoa_to_sheet(orders);
    workbook.Sheets['Orders'] = newWorksheet;
    saveWorkbook();
};

const updateOrder = (order, rowIndex) => {
    const worksheet = getOrdersSheet();
    const orders = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    orders[rowIndex] = [
        order.id, order.name, order.address, order.contact,
        order.dateOfCreation, order.orderDetails, order.totalAmount,
        order.advanceAmount, order.challanDetail, order.orderCompletionStatus
    ];
    const newWorksheet = xlsx.utils.aoa_to_sheet(orders);
    workbook.Sheets['Orders'] = newWorksheet;
    saveWorkbook();
};

const deleteOrder = (rowIndex) => {
    const worksheet = getOrdersSheet();
    const orders = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    orders.splice(rowIndex, 1);
    const newWorksheet = xlsx.utils.aoa_to_sheet(orders);
    workbook.Sheets['Orders'] = newWorksheet;
    saveWorkbook();
};

// Routes

// Get all orders
app.get('/orders', (req, res) => {
    try {
        const orders = readOrders();
        res.json(orders);
    } catch (err) {
        res.status(500).send(err);
    }
});

// Get order by ID
app.get('/orders/:id', (req, res) => {
    try {
        const orders = readOrders();
        const order = orders.find(o => o[0] === req.params.id);
        if (!order) return res.status(404).send('Order not found');
        res.json(order);
    } catch (err) {
        res.status(500).send(err);
    }
});

// Create new order
app.post('/orders', (req, res) => {
    const newOrder = { id: `${Date.now()}`, ...req.body };
    try {
        appendOrder(newOrder);
        res.status(201).json(newOrder);
    } catch (err) {
        res.status(400).send(err);
    }
});

// Update order by ID
app.put('/orders/:id', (req, res) => {
    try {
        const orders = readOrders();
        const orderIndex = orders.findIndex(o => o[0] === req.params.id);
        if (orderIndex === -1) return res.status(404).send('Order not found');
        const updatedOrder = { ...orders[orderIndex], ...req.body };
        updateOrder(updatedOrder, orderIndex + 1); // +1 to account for 0-based index
        res.json(updatedOrder);
    } catch (err) {
        res.status(400).send(err);
    }
});

// Delete order by ID
app.delete('/orders/:id', (req, res) => {
    try {
        const orders = readOrders();
        const orderIndex = orders.findIndex(o => o[0] === req.params.id);
        if (orderIndex === -1) return res.status(404).send('Order not found');
        deleteOrder(orderIndex + 1); // +1 to account for 0-based index
        res.json({ success: true });
    } catch (err) {
        res.status(500).send(err);
    }
});

app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});
