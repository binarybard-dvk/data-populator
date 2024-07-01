import {
  ActivityIcon,
  AlignEndVerticalIcon,
  CombineIcon,
  MergeIcon,
  MousePointerClickIcon,
  SettingsIcon,
  SquareActivityIcon,
  WebhookIcon,
} from "lucide-react";

export default function Features() {
  return (
    <section className="w-full py-12 md:py-24 lg:py-32 xl:py-48 bg-black">
      <div className="container px-4 md:px-6">
        <div className="grid gap-6 items-center">
          <div className="flex flex-col justify-center space-y-8 text-center">
            <div className="space-y-2">
              <h1 className="text-3xl font-bold tracking-tighter sm:text-5xl xl:text-6xl/none bg-clip-text text-transparent bg-gradient-to-r from-white to-gray-500">
                Discover Our Unique Features
              </h1>
              <p className="max-w-[600px] text-zinc-200 md:text-xl dark:text-zinc-100 mx-auto">
                Our features are designed to enhance your productivity and
                streamline your workflow.
              </p>
            </div>
            <div className="w-full max-w-full space-y-4 mx-auto">
              <div className="grid grid-cols-3 gap-8">
                <div className="flex flex-col items-center space-y-2 border-gray-800 p-4 rounded-lg">
                  <div className="p-2 bg-black bg-opacity-50 rounded-full">
                    <MergeIcon className="text-white h-6 w-6 mb-2 opacity-75" />
                  </div>
                  <h2 className="text-xl font-bold text-white">
                    Seamless Integration
                  </h2>
                  <p className="text-zinc-200 dark:text-zinc-100">
                    Effortlessly merge JSON data with Excel templates for quick
                    and easy data population.
                  </p>
                </div>

                <div className="flex flex-col items-center space-y-2 border-gray-800 p-4 rounded-lg">
                  <div className="p-2 bg-black bg-opacity-50 rounded-full">
                    <CombineIcon className="text-white h-6 w-6 mb-2 opacity-75" />
                  </div>
                  <h2 className="text-xl font-bold text-white">
                    Time-Saving Efficiency:
                  </h2>
                  <p className="text-zinc-200 dark:text-zinc-100">
                    Automate repetitive data entry tasks, significantly boosting
                    your productivity.
                  </p>
                </div>

                <div className="flex flex-col items-center space-y-2 border-gray-800 p-4 rounded-lg">
                  <div className="p-2 bg-black bg-opacity-50 rounded-full">
                    <MousePointerClickIcon className="text-white h-6 w-6 mb-2 opacity-75" />
                  </div>
                  <h2 className="text-xl font-bold text-white">
                    High Accuracy
                  </h2>
                  <p className="text-zinc-200 dark:text-zinc-100">
                    Ensure precise data population, eliminating the risk of
                    manual errors.
                  </p>
                </div>
                <div className="flex flex-col items-center space-y-2 border-gray-800 p-4 rounded-lg">
                  <div className="p-2 bg-black bg-opacity-50 rounded-full">
                    <SettingsIcon className="text-white h-6 w-6 mb-2 opacity-75" />
                  </div>
                  <h2 className="text-xl font-bold text-white">
                    Customizable Templates
                  </h2>
                  <p className="text-zinc-200 dark:text-zinc-100">
                    Easily adapt to various Excel template formats and layouts
                    to suit your specific needs.
                  </p>
                </div>
                <div className="flex flex-col items-center space-y-2 border-gray-800 p-4 rounded-lg">
                  <div className="p-2 bg-black bg-opacity-50 rounded-full">
                    <ActivityIcon className="text-white h-6 w-6 mb-2 opacity-75" />
                  </div>
                  <h2 className="text-xl font-bold text-white">
                    Real-Time Processing
                  </h2>
                  <p className="text-zinc-200 dark:text-zinc-100">
                    Experience instant data merging and file generation for
                    immediate results.
                  </p>
                </div>
                <div className="flex flex-col items-center space-y-2 border-gray-800 p-4 rounded-lg">
                  <div className="p-2 bg-black bg-opacity-50 rounded-full">
                    <WebhookIcon className="text-white h-6 w-6 mb-2 opacity-75" />
                  </div>
                  <h2 className="text-xl font-bold text-white">
                    API Availability
                  </h2>
                  <p className="text-zinc-200 dark:text-zinc-100">
                    Use our robust API to integrate DataCraft into your own
                    applications, allowing you to generate populated Excel files
                    programmatically.
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}
