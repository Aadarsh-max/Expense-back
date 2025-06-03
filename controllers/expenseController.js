const xlsx = require('xlsx');
const Expense = require("../models/Expense");

//Add Expense
exports.addExpense = async (req, res) => {
    const userId = req.user.id;

    try{
        const { icon, category, amount, date } = req.body;

        if(!category || !amount || !date){
            return res.status(400).json({ message: "All fields are required" });
        }

        const newExpense = new Expense({
            userId,
            icon,
            category,
            amount,
            date: new Date(date)
        });

        await newExpense.save();
        res.status(200).json(newExpense);
    }catch(error){
        res.status(500).json({ message: "Server Error" });
    }
};

//Get All Expense
exports.getAllExpense = async(req, res) => {
    const userId = req.user.id;

    try{
        const expense = await Expense.find({userId}).sort({date: -1});
        res.json(expense);
    }catch(error){
        res.status(500).json({message: "Server Error"});
    }
};

//Delete Expense
exports.deleteExpense = async(req, res) => {
    try{
        await Expense.findByIdAndDelete(req.params.id);
        res.json({message: "Expense deleted successfully"});
    }catch(error){
        res.status(500).json({message: "Server Error"});
    }
};

//Download Expense Excel
exports.downloadExpenseExcel = async(req, res) => {
    const userId = req.user.id;
    try{
        const expense = await Expense.find({ userId }).sort({ date:-1 });

        const data = expense.map((item) => ({
            Category: item.category,
            Amount: item.amount,
            Date: new Date(item.date).toLocaleDateString('en-CA'), 
        }));

        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(data);
        
       
        ws['!cols'] = [
            { width: 20 }, 
            { width: 15 }, 
            { width: 15 }  
        ];
        
      
        const range = xlsx.utils.decode_range(ws['!ref']);
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const cellAddress = xlsx.utils.encode_cell({ r: row, c: 2 }); 
            if (ws[cellAddress]) {
                ws[cellAddress].t = 's'; 
            }
        }
        
        xlsx.utils.book_append_sheet(wb, ws, "Expense");
        xlsx.writeFile(wb, 'expense_details.xlsx');
        res.download('expense_details.xlsx');
    }catch(error) {
        res.status(500).json({ message: "Server Error" });
    }
};