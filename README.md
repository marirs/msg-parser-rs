Outlook Email Messsage (.msg) parser.
=======================================

A simple parser and reader to deserialize a given Outlook Email Message (.msg) File.  

### Usage
Add this to your `Cargo.toml` file:
```toml
[dependencies]
msg_parser = "0.1.0"
```

### Example

```rust

```

### Requirements
- Rust 1.9+

### Running the given example
```bash
$ cargo run --example parse-email
   Compiling msg_parser v0.1.0 (msg-parser)
    Finished dev [optimized + debuginfo] target(s) in 3.66s
     Running `target/debug/examples/parse-email`
Outlook {
    headers: TransportHeaders {
        content_type: "",
        date: "",
        message_id: "",
        reply_to: "",
    },
    sender: Person {
        name: "",
        email: "",
    },
    ...<clip>
}
```

### Running tests
```bash
cargo t --verbose
```

### Building release
```bash
cargo b --release
```

### Contributions

Fee free to make pull requests to contribute/enhance/add more features/bug fixes.

---
License: MIT