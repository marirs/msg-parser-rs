use msg_parser::Outlook;

fn main() {
    let path = std::env::args()
        .nth(1)
        .unwrap_or("data/test_email.msg".into());
    let outlook = Outlook::from_path(&path).unwrap();

    // Sender & subject
    println!(
        "From:    {} <{}>",
        outlook.sender.name, outlook.sender.email
    );
    println!("Subject: {}", outlook.subject);
    println!("Class:   {}", outlook.message_class);

    // Recipients
    for p in &outlook.to {
        println!("To:      {} <{}>", p.name, p.email);
    }
    for p in &outlook.cc {
        println!("CC:      {} <{}>", p.name, p.email);
    }
    for p in &outlook.bcc {
        println!("BCC:     {} <{}>", p.name, p.email);
    }

    // Dates
    if !outlook.message_delivery_time.is_empty() {
        println!("Delivered: {}", outlook.message_delivery_time);
    }
    if !outlook.client_submit_time.is_empty() {
        println!("Submitted: {}", outlook.client_submit_time);
    }

    // Body preview
    let preview: String = outlook.body.chars().take(100).collect();
    println!("Body:    {}...", preview.replace("\r\n", " "));

    // Attachments
    println!("Attachments: {}", outlook.attachments.len());
    for attach in &outlook.attachments {
        let name = if attach.long_file_name.is_empty() {
            &attach.file_name
        } else {
            &attach.long_file_name
        };
        let method = match attach.attach_method {
            1 => "file",
            5 => "embedded .msg",
            6 => "OLE object",
            _ => "unknown",
        };
        println!(
            "  - {} ({}, {} bytes, {})",
            name,
            attach.extension,
            attach.payload_bytes.len(),
            method,
        );
    }

    // JSON output
    println!("\n--- JSON ---");
    println!("{}", outlook.to_json().unwrap());
}
