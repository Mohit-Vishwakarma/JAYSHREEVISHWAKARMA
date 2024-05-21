const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const winston = require('winston');
const moment = require('moment');
const cors = require('cors');  // Import the cors module

const app = express();
const port = 3000;

// Configure Winston logger
const logger = winston.createLogger({
    level: 'info',
    format: winston.format.combine(
        winston.format.timestamp({
            format: () => moment().format('YYYY-MM-DD HH:mm:ss')
        }),
        winston.format.printf(({ timestamp, level, message }) => `${timestamp} ${level}: ${message}`)
    ),
    transports: [
        new winston.transports.Console(),
        new winston.transports.File({ filename: 'app.log' })
    ]
});

// Middleware
app.use(cors({
    origin: 'http://localhost:3001'  // Replace with your frontend URL
}));
  // Use the cors middleware
app.use(bodyParser.json());

// Load or create Excel file
const excelFilePath = 'orders.xlsx';

let workbook;
if (fs.existsSync(excelFilePath)) {
    logger.info('Loading existing Excel file.');
    workbook = xlsx.readFile(excelFilePath);
} else {
    logger.info('Creating new Excel file.');
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
    logger.info('Excel file saved.');
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
        moment(order.dateOfCreation).format('YYYY-MM-DD HH:mm:ss'), 
        order.orderDetails, order.totalAmount,
        order.advanceAmount, order.challanDetail, order.orderCompletionStatus
    ]);
    const newWorksheet = xlsx.utils.aoa_to_sheet(orders);
    workbook.Sheets['Orders'] = newWorksheet;
    saveWorkbook();
    logger.info(`Order with ID ${order.id} appended.`);
};

const updateOrder = (order, rowIndex) => {
    const worksheet = getOrdersSheet();
    const orders = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    orders[rowIndex] = [
        order.id, order.name, order.address, order.contact,
        moment(order.dateOfCreation).format('YYYY-MM-DD HH:mm:ss'),
        order.orderDetails, order.totalAmount,
        order.advanceAmount, order.challanDetail, order.orderCompletionStatus
    ];
    const newWorksheet = xlsx.utils.aoa_to_sheet(orders);
    workbook.Sheets['Orders'] = newWorksheet;
    saveWorkbook();
    logger.info(`Order with ID ${order.id} updated at row ${rowIndex}.`);
};

const deleteOrder = (rowIndex) => {
    const worksheet = getOrdersSheet();
    const orders = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    const deletedOrder = orders.splice(rowIndex, 1);
    const newWorksheet = xlsx.utils.aoa_to_sheet(orders);
    workbook.Sheets['Orders'] = newWorksheet;
    saveWorkbook();
    logger.info(`Order deleted from row ${rowIndex}.`);
};

// Routes

// Get all orders
app.get('/orders', (req, res) => {
    try {
        const orders = readOrders();
        logger.info('Fetched all orders.');
        res.json(orders);
    } catch (err) {
        logger.error('Error fetching orders: ', err);
        res.status(500).send(err);
    }
});

// Get order by ID
app.get('/orders/:id', (req, res) => {
    try {
        const orders = readOrders();
        const order = orders.find(o => o[0] === req.params.id);
        if (!order) {
            logger.warn(`Order with ID ${req.params.id} not found.`);
            return res.status(404).send('Order not found');
        }
        logger.info(`Fetched order with ID ${req.params.id}.`);
        res.json(order);
    } catch (err) {
        logger.error('Error fetching order: ', err);
        res.status(500).send(err);
    }
});

// Create new order
app.post('/orders', (req, res) => {
    const newOrder = { id: `${Date.now()}`, ...req.body, dateOfCreation: moment().format('YYYY-MM-DD HH:mm:ss') };
    try {
        appendOrder(newOrder);
        logger.info(`Created new order with ID ${newOrder.id}.`);
        res.status(201).json(newOrder);
    } catch (err) {
        logger.error('Error creating order: ', err);
        res.status(400).send(err);
    }
});

// Update order by ID
app.put('/orders/:id', (req, res) => {
    try {
        const orders = readOrders();
        const orderIndex = orders.findIndex(o => o[0] === req.params.id);
        if (orderIndex === -1) {
            logger.warn(`Order with ID ${req.params.id} not found for update.`);
            return res.status(404).send('Order not found');
        }
        const updatedOrder = { ...orders[orderIndex], ...req.body };
        updateOrder(updatedOrder, orderIndex + 1); // +1 to account for 0-based index
        logger.info(`Updated order with ID ${req.params.id}.`);
        res.json(updatedOrder);
    } catch (err) {
        logger.error('Error updating order: ', err);
        res.status(400).send(err);
    }
});

// Delete order by ID
app.delete('/orders/:id', (req, res) => {
    try {
        const orders = readOrders();
        const orderIndex = orders.findIndex(o => o[0] === req.params.id);
        if (orderIndex === -1) {
            logger.warn(`Order with ID ${req.params.id} not found for deletion.`);
            return res.status(404).send('Order not found');
        }
        deleteOrder(orderIndex + 1); // +1 to account for 0-based index
        logger.info(`Deleted order with ID ${req.params.id}.`);
        res.json({ success: true });
    } catch (err) {
        logger.error('Error deleting order: ', err);
        res.status(500).send(err);
    }
});

app.listen(port, () => {
    logger.info(`Server running on http://localhost:${port}`);
});
