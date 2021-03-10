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


            // PRINT WELCOME MESSAGE AND ASK FOR INPUT
            Console.WriteLine("Welcome to your movie collection Egan family! ");
            Console.WriteLine("\n");
            Console.WriteLine("Would you like to search in the Cabin, Latham, the Digital list, or add a new movie? ");

            // GET THE INPUT
            searchLocation = Console.ReadLine();

            // TEST FOR "CABIN" INPUT
            if(searchLocation == "Cabin")
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
            
            // TEST FOR "LATHAM" INPUT
            else if(searchLocation == "Latham")
            {
                // PRINTS A NICE SPACE
                Console.WriteLine();
                
                // BRING IN THE FILE AND APPROPRIATE WORKSHEET
                WorkBook lathamList = new WorkBook("MovieList.xlsx");
                WorkSheet lathamSheet = lathamList.GetWorkSheet("Latham");
                
                // FORMAT THE WORKSHEET

                // PRINT THE WORKSHEET
                Console.WriteLine(lathamSheet);
                Console.WriteLine("\n");
                
            }

            // TEST FOR "DIGITAL" INPUT
            else if (searchLocation == "Digital")
            {
                // PRINTS A NICE SPACE
                Console.WriteLine();

                // BRING IN THE FILE AND WORKSHEET
                WorkBook digitalList = new WorkBook("MovieList.xlsx");
                WorkSheet digitalSheet = digitalList.GetWorkSheet("Digital");

                // FORMAT THE WORKSHEET

                // PRINT THE WORKSHEET
                Console.WriteLine(digitalSheet);                
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
                Console.WriteLine("Would you like to search in the Cabin, Latham, Digital, or add a new movie? ");

                // GET THE INPUT
                searchLocation = Console.ReadLine();                
            }

            

            //ASK FOR ANOTHER INPUT
            Console.WriteLine("Would you like to search somewhere else? (Y / N) ");
            Console.WriteLine();
            newSearch = Console.ReadLine();

            if (newSearch == "Yes")
            {
                Console.WriteLine("Would you like to search in the Cabin, Latham, Digital, or add a new movie? ");
                searchLocation = Console.ReadLine();

                
            }

            else if (newSearch == "No")
            {
                Console.WriteLine("Have a nice day!! Please press enter to exit the program.");
                Console.ReadLine();
            }
        }
    }
}