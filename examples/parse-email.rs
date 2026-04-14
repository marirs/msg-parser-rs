use msg_parser::Outlook;

fn main() {
    let path = std::env::args()
        .nth(1)
        .unwrap_or("data/test_email.msg".into());
    let outlook = Outlook::from_path(&path).unwrap();

    // Display summary (uses the Display impl)
    println!("{}", outlook);

    // Message class and metadata
    println!("Class:       {}", outlook.message_class);
    println!("Importance:  {}", outlook.importance);
    println!("Sensitivity: {}", outlook.sensitivity);

    // Dates
    if !outlook.message_delivery_time.is_empty() {
        println!("Delivered:   {}", outlook.message_delivery_time);
    }
    if !outlook.client_submit_time.is_empty() {
        println!("Submitted:   {}", outlook.client_submit_time);
    }

    // Body preview
    let preview: String = outlook.body.chars().take(100).collect();
    println!("Body:        {}...", preview.replace("\r\n", " "));

    // HTML body (direct or extracted from RTF)
    let html = if !outlook.html.is_empty() {
        Some(outlook.html.clone())
    } else {
        outlook.html_from_rtf()
    };
    if let Some(ref h) = html {
        let preview: String = h.chars().take(80).collect();
        println!("HTML:        {}...", preview);
    }

    // RTF decompression
    if let Some(rtf) = outlook.rtf_decompressed() {
        println!("RTF:         {} bytes decompressed", rtf.len());
    }

    // Attachments
    println!("\nAttachments: {}", outlook.attachments.len());
    for attach in &outlook.attachments {
        // Display impl shows: name (method, size)
        println!("  - {}", attach);

        // Show content-id for inline images
        if !attach.content_id.is_empty() {
            println!("    Content-ID: {}", attach.content_id);
        }

        // Recursively parse embedded .msg attachments
        if let Some(result) = attach.as_message() {
            match result {
                Ok(nested) => {
                    println!("    Embedded message:");
                    println!("      Subject: {}", nested.subject);
                    println!("      From:    {}", nested.sender);
                    println!(
                        "      To:      {:?}",
                        nested.to.iter().map(|p| p.to_string()).collect::<Vec<_>>()
                    );
                }
                Err(e) => println!("    Failed to parse embedded message: {}", e),
            }
        }
    }
}
