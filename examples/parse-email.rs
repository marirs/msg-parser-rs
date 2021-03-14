use msg_parser::Outlook;

fn main() {
    // Create Outlook object
    let outlook = Outlook::from_path("data/test_email_4.msg").unwrap();

    // Flush as json string
    let _json_string = outlook.to_json();

    println!("{:#?}", outlook);
}
