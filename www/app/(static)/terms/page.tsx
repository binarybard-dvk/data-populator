export default function Terms() {
  return (
    <div className="py-12">
      <h2 className="text-2xl font-bold mb-4">Terms and Conditions</h2>

      <p className="mb-4">Welcome to DataCraft!</p>

      <p className="mb-4">
        By accessing and using DataCraft, you agree to comply with and be bound
        by the following terms and conditions of use. Please read these terms
        carefully before using our service.
      </p>

      <ol className="list-decimal pl-6 mb-4">
        <li className="mb-2">
          <strong className="font-bold">Use of Service:</strong> DataCraft
          allows users to merge JSON data with Excel templates to generate
          populated Excel files. You agree to use the service solely for lawful
          purposes and in compliance with all applicable laws and regulations.
        </li>
        <li className="mb-2">
          <strong className="font-bold">Data Security:</strong> We take data
          security seriously. While using DataCraft, you are responsible for
          maintaining the confidentiality of your account credentials and
          ensuring the security of your data.
        </li>
        <li className="mb-2">
          <strong className="font-bold">API Usage:</strong> Our API is available
          for integration into your own applications, subject to adherence to
          our API usage policies and guidelines.
        </li>
        <li className="mb-2">
          <strong className="font-bold">Intellectual Property:</strong> All
          content and functionality on DataCraft, including but not limited to
          text, graphics, logos, and images, are the property of DataCraft. You
          agree not to reproduce, distribute, or create derivative works from
          any content without our express permission.
        </li>
        <li className="mb-2">
          <strong className="font-bold">Disclaimer of Warranties:</strong>{" "}
          DataCraft is provided on an "as is" and "as available" basis. We make
          no warranties, express or implied, regarding the service's
          reliability, availability, or suitability for your specific needs.
        </li>
        <li className="mb-2">
          <strong className="font-bold">Limitation of Liability:</strong> In no
          event shall DataCraft or its affiliates be liable for any direct,
          indirect, incidental, special, or consequential damages arising out of
          or in any way connected with the use or inability to use our service.
        </li>
        <li className="mb-2">
          <strong className="font-bold">Changes to Terms:</strong> DataCraft
          reserves the right to update or modify these terms and conditions at
          any time without prior notice. It is your responsibility to review
          these terms periodically for changes.
        </li>
        <li className="mb-2">
          <strong className="font-bold">Governing Law:</strong> These terms and
          conditions shall be governed by and construed in accordance with the
          laws of [Your Jurisdiction], without regard to its conflict of law
          provisions.
        </li>
      </ol>

      <p className="mb-4">
        By using DataCraft, you acknowledge that you have read, understood, and
        agree to be bound by these terms and conditions. If you do not agree
        with any part of these terms, please do not use our service.
      </p>
    </div>
  );
}
