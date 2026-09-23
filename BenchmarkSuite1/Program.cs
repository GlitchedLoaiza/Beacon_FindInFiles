// -----------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// -----------------------------------------------------------------------

using BenchmarkDotNet.Running;

namespace BenchmarkSuite1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 1 && args[0] == "--prepare-production-scaling")
            {
                var corpus = ProductionSearchCorpus.Prepare();
                System.Console.WriteLine("Corpus: " + corpus.Root);
                System.Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(corpus.Manifest,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
                return;
            }

            var _ = BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        }
    }
}
