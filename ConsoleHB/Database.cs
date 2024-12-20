﻿/*
 * Author: David Beltran
 */

using MySql.Data.MySqlClient;
using System;
using System.Collections;

namespace ConsoleHB
{
    /// <summary>
    /// Handles entering ArrayList to MySQL database
    /// </summary>
    public class Database
    {
        private readonly MySqlConnection conn;
        private readonly string server;
        private readonly string database;
        private readonly string uid;
        private readonly string password;

        /// <summary>
        /// Database class constructor
        /// </summary>
        public Database()
        {
            this.server = "localhost";
            this.database = "expenses";
            this.uid = "root";
            this.password = "Diska1725!";
            string connString = "SERVER=" + this.server + ";DATABASE=" + this.database +
                ";UID=" + this.uid + ";PASSWORD=" + this.password + ";";
            this.conn = new MySqlConnection(connString);
        }

        /// <summary>
        /// Opens MySQL connection
        /// Mostly created to ensure connection was made when initially designed Database class
        /// </summary>
        private void OpenConnection()
        {
            try
            {
                this.conn.Open();
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Creates and fills 'Transactions' table in MySQL database
        /// </summary>
        /// <param name="expenses">ArrayList made from pdf file</param>
        public void CreateTable(ArrayList expenses)
        {
            OpenConnection();
            string sqlTable = "CREATE TABLE Transactions (" +
                "Date varchar(25)," +
                "Details text," +
                "Amount float," +
                "Category text," +
                "Subcategory text)";
            MySqlCommand cmd = new MySqlCommand(sqlTable, this.conn);
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (MySqlException ex) // catches if table already exists
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                FillTable(expenses);
            }
        }

        /// <summary>
        /// Fills 'Transactions' table in MySQL database
        /// </summary>
        /// <param name="expenses">ArrayList made from pdf file</param>
        /// TODO
        /// - Find Python equivalent of ExecuteMany() method.
        /// - Research parameter declerations... May need safer design than Parameters.Clear() within for loop
        /// - Research SQL injection... Does this pass?
        private void FillTable(ArrayList expenses)
        {
            string fillTable = "INSERT INTO Transactions(Date, Details, Amount, Category, Subcategory) " +
                "VALUES (@Date, @Details, @Amount, @Category, @Subcategory)";
            MySqlCommand cmd = new MySqlCommand(fillTable, this.conn);
            
            foreach (Expense exp in expenses)
            {
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@Date", exp.Date);
                cmd.Parameters.AddWithValue("@Details", exp.Detail);
                cmd.Parameters.AddWithValue("@Amount", exp.Amount);
                cmd.Parameters.AddWithValue("@Category", exp.Category.ToString());
                cmd.Parameters.AddWithValue("@Subcategory", exp.SubCategory.ToString());
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Close connection to MySQL database for each Database class instance
        /// </summary>
        public void CloseDatabase()
        {
            this.conn.Close();
        }
    }
}