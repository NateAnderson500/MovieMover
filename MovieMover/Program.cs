using System;
using IronXL;

namespace MovieMover
{
    class Program
    {
        static void Main(string[] args)
        {
            // DECALRE GLOBAL VARIABLES
            string searchLocation = "";
            string newSearch = "";
            Program input = new Program();

            // PRINT WELCOME MESSAGE AND ASK FOR INPUT
            Console.WriteLine("Welcome to your movie collection! ");
            Console.WriteLine("\n");
            Console.WriteLine("Would you like to search in the Cabin, Home, the Digital list, or add a new movie? ");

            // GET THE INPUT
            searchLocation = Console.ReadLine();

            // TEST FOR "CABIN" INPUT
            if(searchLocation == "Cabin")
            {
                input.cabin();
            }
            
            // TEST FOR "LATHAM" INPUT
            else if(searchLocation == "Home")
            {
                input.home();
            }

            // TEST FOR "DIGITAL" INPUT
            else if (searchLocation == "Digital")
            {
                input.digital();
            }
            

            // TEST FOR "NEW" INPUT
            else if(searchLocation == "New")
            {
                // Methods go here
                Console.WriteLine("Not done yet. Probably because I don't know what I am doing");
                Console.ReadLine();
            }

            // RETRY IF THE INPUT DOES NOT MATCH
            else
            {
                // PRINT ERROR MESSAGE AND TRY AGAIN
                Console.WriteLine("Please enter a valid input. (Try capitalizing the first letter) ");
                Console.WriteLine("\n");
                Console.WriteLine("Would you like to search in the Cabin, Home, Digital, or add a new movie? ");

                // GET THE INPUT
                searchLocation = Console.ReadLine();                
            }

            

            //ASK FOR ANOTHER INPUT
            Console.WriteLine("Would you like to search somewhere else? (Y / N) ");
            Console.WriteLine();
            newSearch = Console.ReadLine();

            if (newSearch == "Yes")
            {
                Console.WriteLine("Would you like to search in the Cabin, Home, Digital, or add a new movie? ");
                searchLocation = Console.ReadLine();

                
            }

            else if (newSearch == "No")
            {
                Console.WriteLine("Have a nice day!! Please press enter to exit the program.");
                Console.ReadLine();
            }
        }

        // METHOD TO PRINT THE CABIN LIST
        public void cabin()
        {
            // PRINTS A NICE SPACE
            Console.WriteLine();

            // BRING IN THE FILE AND APPROPRIATE WORKSHEET
            WorkBook CabinList = new WorkBook("MovieList.xlsx");
            WorkSheet cabinSheet = CabinList.GetWorkSheet("Cabin");

            // FORMAT THE WORKSHEET

            // PRINT THE WORKSHEET
            Console.WriteLine(cabinSheet);
            Console.WriteLine("\n");
        }

        public void home()
        {
            // PRINTS A NICE SPACE
            Console.WriteLine();

            // BRING IN THE FILE AND APPROPRIATE WORKSHEET
            var digitalBook = new WorkBook("MovieList.xlsx");
            var digitalSheet = digitalBook.GetWorkSheet("Home");
            var titleRange = digitalSheet.GetRange("A1:A203");
            var locationRange = digitalSheet.GetRange("B1:B199");
            var typeRange = digitalSheet.GetRange("C1:C199");
            var conditionRange = digitalSheet.GetRange("D1:D199");
            var digitalRange = digitalSheet.GetRange("E1:E199");


            Console.Write(titleRange + "\t" + locationRange);
        }

        public void digital()
        {
            // PRINTS A NICE SPACE
            Console.WriteLine();

            // BRING IN THE FILE AND WORKSHEET
            var digitalBook = new WorkBook("MovieList.xlsx");
            var digitalSheet = digitalBook.GetWorkSheet("Digital");
            var titleRange = digitalSheet.GetRange("A1:A100");
            foreach (var cellA in titleRange)
            {
                Console.WriteLine(cellA.Value);
            }
        }
    }
}