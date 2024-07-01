const express = require('express');
const cors = require('cors');
const app = express();
require('dotenv').config();
const { MongoClient, ServerApiVersion, ObjectId } = require('mongodb');
const XLSX = require('xlsx');
const fs = require('fs');

const port = process.env.PORT || 5000;

const cloudinary = require('cloudinary').v2;

// Configure Cloudinary with your credentials
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

// Middleware
app.use(cors());
app.use(express.json());

// Generate Cloudinary signature route
app.post('/generate-signature', (req, res) => {
  const timestamp = Math.round(new Date().getTime() / 1000);

  const signature = cloudinary.utils.api_sign_request(
    {
      timestamp: timestamp,
      upload_preset: 'your_upload_preset',
    },
    process.env.CLOUDINARY_API_SECRET
  );

  res.json({ timestamp, signature });
});

const uri = `mongodb+srv://${process.env.DB_USER}:${process.env.DB_PASS}@cluster0.v2ftvhq.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0`;

// Create a MongoClient with a MongoClientOptions object to set the Stable API version
const client = new MongoClient(uri, {
  serverApi: {
    version: ServerApiVersion.v1,
    strict: true,
    deprecationErrors: true,
  }
});

// Function to write order data to Excel
const writeOrderToExcel = async (order) => {
  const filePath = './orders.xlsx';

  // Prepare order data, removing the status field
  const orderData = {
    Name: order.name,
    Email: order.email,
    Address: order.address,
    Contact: order.number,
    Pin: order.pin,
    ProductName: order.productName,
    Quantity: order.quantity,
    TotalPrice: order.totalPrice
  };

  // Check if the file exists
  let workbook;
  if (fs.existsSync(filePath)) {
    // Read existing file
    workbook = XLSX.readFile(filePath);
  } else {
    // Create a new workbook and worksheet
    workbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet([]);
    workbook.Sheets['Orders'] = newWorksheet;

    // Set the headers bold
    const headers = Object.keys(orderData);
    const headerRow = headers.map(header => ({ v: header, s: { font: { bold: true } } }));
    XLSX.utils.sheet_add_aoa(newWorksheet, [headerRow], { origin: 'A1' });

    // Adjust column widths
    newWorksheet['!cols'] = headers.map((header, index) => {
      if (['Name', 'Email', 'Address'].includes(header)) {
        return { wch: 30 };
      }
      return { wch: 15 };
    });
  }

  const worksheet = workbook.Sheets['Orders'];

  // Convert the sheet to JSON and add the new order
  const data = XLSX.utils.sheet_to_json(worksheet);
  data.push(orderData);

  // Write the updated data back to the worksheet
  const newWorksheet = XLSX.utils.json_to_sheet(data);
  workbook.Sheets['Orders'] = newWorksheet;

  // Adjust column widths again to make sure new columns are adjusted
  const headers = Object.keys(orderData);
  newWorksheet['!cols'] = headers.map((header, index) => {
    if (['Name', 'Email', 'Address'].includes(header)) {
      return { wch: 30 };
    }
    return { wch: 15 };
  });

  // Write the workbook to a file
  XLSX.writeFile(workbook, filePath);
};

async function run() {
  try {
    console.log("Connecting to MongoDB...");
    await client.connect();
    console.log("Connected to MongoDB");

    const productsCollection = client.db('sapiensCare').collection('products');
    const bookingProductsCollection = client.db('sapiensCare').collection('bookingProducts');

    app.get('/products', async(req, res) =>{
      try {
        const cursor = productsCollection.find();
        const result = await cursor.toArray();
        res.send(result);
      } catch (error) {
        console.error('Error fetching products:', error);
        res.status(500).send({ message: 'Internal server error' });
      }
    });

    app.post('/products', async(req, res) =>{
      try {
        const product = req.body;
        const result = await productsCollection.insertOne(product);
        res.send(result);
      } catch (error) {
        console.error('Error adding product:', error);
        res.status(500).send({ message: 'Internal server error' });
      }
    });

    app.get('/products/:id', async (req, res) => {
      try {
        const id = req.params.id;
        const query = { _id: new ObjectId(id) };
        const product = await productsCollection.findOne(query);
        res.send(product);
      } catch (error) {
        console.error('Error fetching product by ID:', error);
        res.status(500).send({ message: 'Internal server error' });
      }
    });

    app.get('/bookingProducts', async(req, res) => {
      try {
        const query = {};
        const bookingProducts = await bookingProductsCollection.find(query).toArray();
        res.send(bookingProducts);
      } catch (error) {
        console.error('Error fetching booking products:', error);
        res.status(500).send({ message: 'Internal server error' });
      }
    });

    app.post('/bookingProducts', async(req, res) => {
      try {
        const bookingProduct = req.body;
        const result = await bookingProductsCollection.insertOne(bookingProduct);
        res.send(result);
      } catch (error) {
        console.error('Error adding booking product:', error);
        res.status(500).send({ message: 'Internal server error' });
      }
    });

    app.patch('/bookingProducts/:id', async (req, res) => {
      try {
        const id = req.params.id;
        const status = req.body.status;
        const query = { _id: new ObjectId(id) };
        const update = { $set: { status: status } };
        
        const result = await bookingProductsCollection.updateOne(query, update);
        if (result.modifiedCount === 1) {
          // Fetch the updated order details
          const updatedOrder = await bookingProductsCollection.findOne(query);
          if (status === "confirmed") {
            // Write the confirmed order to Excel
            await writeOrderToExcel(updatedOrder);
          }
          res.status(200).json({ message: 'Order status updated successfully.' });
        } else {
          res.status(404).json({ message: 'Order not found.' });
        }
      } catch (error) {
        console.error('Error updating order status:', error);
        res.status(500).json({ message: 'Internal server error' });
      }
    });

    app.delete('/bookingProducts/:id', async (req, res) => {
      try {
        const id = req.params.id;
        const query = { _id: new ObjectId(id) };
    
        const deleteResult = await bookingProductsCollection.deleteOne(query);
    
        if (deleteResult.deletedCount === 1) {
          res.status(200).json({ message: 'Order deleted successfully.' });
        } else {
          res.status(404).json({ message: 'Order not found.' });
        }
      } catch (error) {
        console.error('Error deleting order:', error);
        res.status(500).json({ message: 'Internal server error' });
      }
    });

    // Endpoint to download the orders Excel file
    app.get('/download-orders', (req, res) => {
      const filePath = './orders.xlsx';
      if (fs.existsSync(filePath)) {
        res.download(filePath);
      } else {
        res.status(404).send('No orders file found');
      }
    });

    // Send a ping to confirm a successful connection
    await client.db("admin").command({ ping: 1 });
    console.log("Pinged your deployment. You successfully connected to MongoDB!");
  } catch (error) {
    console.error("Error running the server:", error);
  }
}
run().catch(console.dir);

app.get('/', (req, res) =>{
    res.send('Sapiens Care is running')
})

app.listen(port, () =>{
    console.log(`Sapiens Care is running on port ${port}`)
})
