import xlsx from 'xlsx';
import Income from '../models/Income.js';

// Add Income Source
export const addIncome = async (req, res) => {
    const userId = req.user.id;

    try {
        const { icon, source, amount, date } = req.body;

        if (!source || !amount || !date) {
            return res.status(400).json({ message: "All fields are required" });
        }

        const newIncome = new Income({
            userId,
            icon,
            source,
            amount,
            date: new Date(date),
        });

        await newIncome.save();
        res.status(200).json(newIncome);
    } catch (error) {
        res.status(500).json({ message: "Server Error" });
    }
};

// Get All Income Source
export const getAllIncome = async (req, res) => {
    const userId = req.user.id;

    try {
        const income = await Income.find({ userId }).sort({ date: -1 });
        res.json(income);
    } catch (error) {
        res.status(500).json({ message: "Server Error" });
    }
};

// Delete Income Source
export const deleteIncome = async (req, res) => {
    try {
        await Income.findByIdAndDelete(req.params.id);
        res.json({ message: "Income deleted successfully" });
    } catch (error) {
        res.status(500).json({ message: "Server Error" });
    }
};

// Download Excel
export const downloadIncomeExcel = async (req, res) => {
    const userId = req.user.id;
    try {
        const income = await Income.find({ userId }).sort({ date: -1 });

        const data = income.map((item) => ({
            Source: item.source,
            Amount: item.amount,
            Date: new Date(item.date).toLocaleDateString('en-CA'),
        }));

        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(data);

        ws['!cols'] = [
            { width: 20 },
            { width: 15 },
            { width: 15 },
        ];

        const range = xlsx.utils.decode_range(ws['!ref']);
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const cellAddress = xlsx.utils.encode_cell({ r: row, c: 2 });
            if (ws[cellAddress]) {
                ws[cellAddress].t = 's';
            }
        }

        xlsx.utils.book_append_sheet(wb, ws, "Income");
        xlsx.writeFile(wb, 'income_details.xlsx');
        res.download('income_details.xlsx');
    } catch (error) {
        res.status(500).json({ message: "Server Error" });
    }
};
