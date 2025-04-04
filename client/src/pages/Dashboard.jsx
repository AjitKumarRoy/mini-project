import React, { useEffect, useState } from "react";
import api from "../api/api";
import { Link } from "react-router-dom";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Trash2 } from "lucide-react"; // Import Trash2 (delete icon)
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";

const Dashboard = () => {
  const [spreadsheets, setSpreadsheets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = useState(false);
  const [spreadsheetToDelete, setSpreadsheetToDelete] = useState(null);

  useEffect(() => {
    fetchSpreadsheets();
  }, []);

  async function fetchSpreadsheets() {
    try {
      const res = await api.get("/api/sheets/listSpreadSheets");
      setSpreadsheets(res.data.spreadsheets || []);
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  }

  const handleDeleteClick = (spreadsheetId) => {
    setSpreadsheetToDelete(spreadsheetId);
    setIsDeleteDialogOpen(true);
  };

  const confirmDelete = async () => {
    setIsDeleteDialogOpen(false);
    setLoading(true);
    try {
      await api.delete(`/api/sheets/${spreadsheetToDelete}/deleteSpreadSheet`);
      fetchSpreadsheets(); // Refresh the list after deletion
    } catch (error) {
      console.error("Error deleting spreadsheet:", error);
      // Optionally display an error message to the user
    } finally {
      setLoading(false);
      setSpreadsheetToDelete(null);
    }
  };

  const cancelDelete = () => {
    setIsDeleteDialogOpen(false);
    setSpreadsheetToDelete(null);
  };

  return (
    <div className="p-4 bg-gray-100 flex flex-col items-center">
      {loading ? (
        <div className="flex justify-center items-center h-40">
          <div className="animate-spin rounded-full h-16 w-16 border-t-2 border-b-2 border-blue-500"></div>
        </div>
      ) : (
        <div className="w-full max-w-md flex flex-col gap-4">
          {/* Custom Template Section */}
          <Card className="bg-white shadow-lg p-4 rounded-xl">
            <CardContent className="flex flex-col items-center text-center">
              <h2 className="text-lg font-semibold text-gray-700 mb-2">
                Create a New Spreadsheet
              </h2>
              <Link to="/create-spreadsheet">
                <Button
                  variant="outline"
                  className="px-6 py-2 text-blue-600 border-blue-500"
                >
                  + New Spreadsheet
                </Button>
              </Link>
            </CardContent>
          </Card>

          {/* List of Spreadsheets */}
          <Card className="bg-white shadow-lg p-4 rounded-xl">
            <CardContent className="pb-2"> {/* Added pb-2 to allow scroll without cutting last item */}
              <h2 className="text-lg font-semibold text-gray-700 mb-3 flex items-center justify-between">
                Total Spreadsheets: {spreadsheets.length}
              </h2>
              {spreadsheets.length === 0 ? (
                <p className="text-gray-500 text-center">
                  No spreadsheets found
                </p>
              ) : (
                <div className="flex flex-col gap-2 max-h-[300px] overflow-y-auto"> {/* Added max-h and overflow-y */}
                  {spreadsheets.map((sheet, index) => (
                    <div
                      key={sheet.id}
                      className="flex items-center justify-between p-3 bg-gray-50 rounded-lg shadow-sm hover:bg-gray-100 transition"
                    >
                      <Link
                        to={`/sheets/${sheet.id}`}
                        className="flex-grow"
                      >
                        {index + 1}. {sheet.name}
                      </Link>
                      <button
                        onClick={() => handleDeleteClick(sheet.id)}
                        className="ml-2 text-red-500 hover:text-red-700 focus:outline-none transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </CardContent>
          </Card>
        </div>
      )}

      {/* Delete Confirmation Dialog */}
      <Dialog open={isDeleteDialogOpen} onOpenChange={setIsDeleteDialogOpen}>
        <DialogContent className="max-w-md absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2">
          <DialogHeader>
            <DialogTitle>Confirm Delete</DialogTitle>
          </DialogHeader>
          <p className="mb-4">
            Are you sure you want to delete this spreadsheet? This action cannot be undone.
          </p>
          <DialogFooter>
            <Button variant="secondary" onClick={cancelDelete}>
              Cancel
            </Button>
            <Button variant="destructive" onClick={confirmDelete}>
              Delete
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
};

export default Dashboard;